#!/usr/bin/env bash
# Reusable Kali/Debian cybersecurity workstation bootstrap.
set -Eeuo pipefail
umask 077

JOBS=2
MAX_ATTEMPTS=3
FAILURES=()
MISSING=()

usage() {
    cat <<'EOF'
Usage: bash makeitmyworkstation.sh
       bash makeitmyworkstation.sh -h

With no arguments, installs the complete workstation. The only option is -h.
The script must be executed as root. It securely configures Kali Rolling first.
Supported: 64-bit Kali, Debian, and Raspberry Pi OS; amd64 and arm64.
It installs large Kali tool and wordlist collections, Docker, current Python,
Go and Rust, all ProjectDiscovery tools, and history-derived CLI tools.
Install failures are retried up to three times with repair steps and reported.
Exit codes: 0 complete; 1 invalid invocation/prerequisite; 2 incomplete install.
EOF
}

die() { printf 'ERROR: %s\n' "$*" >&2; exit 1; }
note() { printf '\n==> %s\n' "$*"; }

on_error() {
    local rc=$?
    printf 'Stopped at line %s (exit %s). Logs: %s\n' "${BASH_LINENO[0]}" "$rc" "${RUN:-not-created}" >&2
    exit "$rc"
}

parse_args() {
    (($# == 0)) && return
    (($# == 1)) && [[ $1 == -h ]] && { usage; exit 0; }
    die 'The only supported option is -h. Run without arguments to install.'
}

detect_platform() {
    [[ -r /etc/os-release ]] || die 'Missing /etc/os-release'
    # shellcheck disable=SC1091
    source /etc/os-release
    case "${ID:-}" in
        kali|debian|raspbian) ;;
        ubuntu) die 'Ubuntu is not supported: mixing Kali and Ubuntu repositories is unsafe' ;;
        *) die "Unsupported distribution: ${ID:-unknown}; use Kali, Debian, or Raspberry Pi OS" ;;
    esac
    ARCH=$(dpkg --print-architecture)
    case "$ARCH" in
        amd64) RUST_TARGET=x86_64-unknown-linux-gnu ;;
        arm64) RUST_TARGET=aarch64-unknown-linux-gnu ;;
        *) die "Unsupported architecture $ARCH. Use a 64-bit OS on Raspberry Pi." ;;
    esac
    ((EUID == 0)) || die 'This script must be executed as root'
    HOME=/root
    BASE="$HOME/.local/share/cyber-workstation"
    BIN="$BASE/bin"
    export CARGO_HOME="$BASE/cargo" RUSTUP_HOME="$BASE/rustup"
    export UV_PYTHON_INSTALL_DIR="$BASE/python"
    export UV_TOOL_DIR="$BASE/python-tools" UV_TOOL_BIN_DIR="$BIN"
    export UV_PYTHON_BIN_DIR="$BIN"
    export GOPATH="$BASE/gopath" GOBIN="$BIN"
    if command -v nproc >/dev/null; then JOBS=$(nproc); ((JOBS > 4)) && JOBS=4; fi
    export GOMAXPROCS="$JOBS" CARGO_BUILD_JOBS="$JOBS"
    export PATH="$BIN:$BASE/go/current/bin:$CARGO_HOME/bin:$BASE/pdtm:$PATH"
}

configure_kali_repository() {
    note 'Configure the official Kali Rolling repository and archive keyring'
    local key_url='https://archive.kali.org/archive-keyring.gpg'
    local keyring='/usr/share/keyrings/kali-archive-keyring.gpg'
    local source_file='/etc/apt/sources.list.d/kali.sources'
    local expected_fingerprint='827C8569F2518CC677FECA1AED65462EC8D5E4C5'
    local downloaded_key="$RUN/kali-archive-keyring.gpg"
    local staged_source="$RUN/kali.sources"

    # Fresh Debian needs these before it can authenticate and use Kali metadata.
    if ! command -v curl >/dev/null || ! command -v gpg >/dev/null ||
        ! dpkg-query -W -f='${Status}' ca-certificates 2>/dev/null | grep -Fq 'install ok installed'; then
        retry 'Refresh the original distribution repositories' apt-get update -o Acquire::Retries=3 || \
            die 'Could not refresh the original repositories to install repository prerequisites'
        apt_install ca-certificates curl gnupg || \
            die 'Could not install ca-certificates, curl, and gnupg'
    fi
    fetch "$key_url" "$downloaded_key"
    gpg --batch --quiet --show-keys --with-colons "$downloaded_key" 2>/dev/null |
        awk -F: '$1 == "fpr" { print $10 }' |
        grep -Fxq "$expected_fingerprint" || \
        die 'Downloaded Kali archive key does not contain the expected official fingerprint'
    install -d -m 0755 /usr/share/keyrings /etc/apt/sources.list.d
    install -m 0644 "$downloaded_key" "$keyring.new"
    mv -f "$keyring.new" "$keyring"

    cat > "$staged_source" <<'EOF'
# See https://www.kali.org/docs/general-use/kali-apt-sources/
Types: deb
URIs: http://http.kali.org/kali/
Suites: kali-rolling
Components: main contrib non-free non-free-firmware
Signed-By: /usr/share/keyrings/kali-archive-keyring.gpg
EOF
    if [[ -f $source_file ]] && ! cmp -s "$staged_source" "$source_file"; then
        cp -p "$source_file" "$RUN/kali.sources.before"
    fi
    install -m 0644 "$staged_source" "$source_file.new"
    mv -f "$source_file.new" "$source_file"

    if ! grep -Fxq 'URIs: http://http.kali.org/kali/' "$source_file" ||
        ! grep -Fxq 'Suites: kali-rolling' "$source_file" ||
        ! grep -Fxq 'Signed-By: /usr/share/keyrings/kali-archive-keyring.gpg' "$source_file"; then
        die 'Kali repository file verification failed'
    fi
    retry 'Verify signed Kali Rolling metadata' apt-get update -o Acquire::Retries=3 || \
        die 'Kali repository metadata could not be authenticated and downloaded'
}

package_plan() {
    REQUIRED=(ca-certificates curl wget git gnupg jq unzip zip xz-utils
        build-essential pkg-config libssl-dev libffi-dev libpcap-dev
        python3 python3-venv python3-dev zsh bash-completion util-linux
        nmap hashcat docker.io nodejs npm)
    OPTIONAL=(kali-linux-headless kali-tools-top10 kali-tools-information-gathering
        kali-tools-vulnerability kali-tools-web kali-tools-database
        kali-tools-passwords kali-tools-exploitation kali-tools-post-exploitation
        kali-tools-forensics kali-tools-reverse-engineering kali-tools-sniffing-spoofing
        kali-tools-social-engineering kali-tools-crypto-stego kali-tools-fuzzing
        kali-tools-identify kali-tools-protect kali-tools-respond kali-tools-reporting
        git-lfs ripgrep fd-find bat fzf tmux htop tree nano vim
        silversearcher-ag shellcheck bats pipx ruby ruby-dev perl php-cli
        clang cmake make libclang-dev libkrb5-dev krb5-user
        dnsutils whois traceroute iputils-ping iproute2 net-tools
        netcat-openbsd socat openssh-client openvpn tcpdump tshark
        sqlite3 libxml2-utils xmlstarlet file libimage-exiftool-perl p7zip-full
        plocate pciutils
        lsof strace gdb dnsrecon sslscan whatweb ffuf gobuster
        john hashcat-data clinfo azure-cli
        feroxbuster masscan massdns nikto sqlmap wafw00f
        dnsenum enum4linux-ng smbclient ldap-utils nfs-common
        impacket-scripts python3-impacket netexec certipy-ad bloodhound-ce-python
        responder evil-winrm hydra exploitdb metasploit-framework
        penelope theharvester urlcrazy unicornscan freerdp3-x11
        reptyr nvtop age tesseract-ocr ffmpeg seclists wordlists)
}

fetch() {
    curl --fail --location --silent --show-error --retry 3 \
        --connect-timeout 20 --max-time 900 --proto '=https' --proto-redir '=https' \
        "$1" -o "$2"
}

retry() {
    local label=$1 attempt rc
    shift
    for ((attempt=1; attempt<=MAX_ATTEMPTS; attempt++)); do
        note "$label (attempt $attempt/$MAX_ATTEMPTS)"
        if (set -Ee; "$@"); then return 0; fi
        rc=$?
        printf 'Attempt %d for %s failed with exit %d.\n' "$attempt" "$label" "$rc" >&2
        ((attempt < MAX_ATTEMPTS)) && sleep $((attempt * 2))
    done
    return "$rc"
}

candidate() {
    local v
    v=$(LC_ALL=C apt-cache policy "$1" | awk '/Candidate:/ {print $2; exit}')
    [[ -n $v && $v != '(none)' ]]
}

apt_install() {
    local attempt rc
    for ((attempt=1; attempt<=MAX_ATTEMPTS; attempt++)); do
        if env DEBIAN_FRONTEND=noninteractive NEEDRESTART_MODE=l \
            apt-get -y --no-remove --no-install-recommends \
            -o Acquire::Retries=3 -o DPkg::Lock::Timeout=120 install "$@"; then
            return 0
        fi
        rc=$?
        printf 'APT install failed (attempt %d/%d, exit %d); running repair steps.\n' \
            "$attempt" "$MAX_ATTEMPTS" "$rc" >&2
        dpkg --configure -a || true
        env DEBIAN_FRONTEND=noninteractive apt-get -f install -y \
            -o Acquire::Retries=3 -o DPkg::Lock::Timeout=120 || true
        apt-get update -o Acquire::Retries=3 || true
    done
    return "$rc"
}

install_packages() {
    note 'Refresh APT metadata and record package state'
    dpkg-query -W -f='${binary:Package}\t${Version}\n' > "$RUN/packages-before.tsv"
    apt-mark showmanual > "$RUN/manual-before.txt"
    retry 'Refresh APT metadata' apt-get update -o Acquire::Retries=3 || \
        die 'APT metadata refresh failed after repair attempts'
    local p
    for p in "${REQUIRED[@]}"; do
        candidate "$p" || die "Required APT package unavailable: $p"
    done
    apt_install "${REQUIRED[@]}"
    # Independent transactions prevent a single unavailable tool from blocking others.
    for p in "${OPTIONAL[@]}"; do
        if candidate "$p"; then
            if ! apt_install "$p"; then FAILURES+=("apt:$p"); fi
        else
            MISSING+=("apt:$p")
        fi
    done
    if candidate docker-compose-v2; then
        apt_install docker-compose-v2 || FAILURES+=(docker-compose-v2)
    elif candidate docker-compose; then
        apt_install docker-compose || FAILURES+=(docker-compose)
    else
        MISSING+=(docker-compose)
    fi
}

install_python() {
    note 'Install uv in an isolated bootstrap environment and resolve stable Python'
    /usr/bin/python3 -m venv "$BASE/uv-bootstrap"
    "$BASE/uv-bootstrap/bin/python" -m pip install --upgrade --index-url https://pypi.org/simple pip uv
    ln -sfn "$BASE/uv-bootstrap/bin/uv" "$BIN/uv"
    # Request "3" explicitly: latest stable CPython, not a default older runtime.
    "$BIN/uv" python install 3
    "$BIN/uv" python install 3.12
    LATEST_PYTHON=$("$BIN/uv" python find --managed-python 3)
    # Explicit name avoids changing python/python3 used by OS or existing scripts.
    ln -sfn "$LATEST_PYTHON" "$BIN/python-latest"
    "$BIN/python-latest" --version
}

install_command_aliases() {
    # Debian renames these commands to avoid package-name collisions.
    [[ -e $BIN/bat || ! -x /usr/bin/batcat ]] || ln -s /usr/bin/batcat "$BIN/bat"
    [[ -e $BIN/fd || ! -x /usr/bin/fdfind ]] || ln -s /usr/bin/fdfind "$BIN/fd"
}

install_go() {
    note 'Resolve and checksum-verify the current stable Go release'
    local filename digest version stage
    fetch 'https://go.dev/dl/?mode=json' "$RUN/go-releases.json"
    read -r filename digest version < <(jq -r --arg arch "$ARCH" \
        '[.[] | select(.stable == true) | .files[] |
          select(.os == "linux" and .arch == $arch and .kind == "archive")][0] |
          [.filename, .sha256, .version] | @tsv' "$RUN/go-releases.json")
    [[ $version =~ ^go[0-9]+\.[0-9]+(\.[0-9]+)?$ ]] || die 'Invalid Go version metadata'
    [[ $filename == "$version.linux-$ARCH.tar.gz" && $digest =~ ^[a-f0-9]{64}$ ]] || die 'Invalid Go archive metadata'
    mkdir -p "$BASE/go"
    if [[ ! -x $BASE/go/$version/bin/go ]]; then
        stage=$(mktemp -d "$BASE/go/stage.XXXXXX")
        fetch "https://go.dev/dl/$filename" "$stage/go.tar.gz"
        printf '%s  %s\n' "$digest" "$stage/go.tar.gz" | sha256sum --check -
        tar -xzf "$stage/go.tar.gz" -C "$stage" --no-same-owner
        [[ ! -e $BASE/go/$version ]] || die "Incomplete existing Go directory: $BASE/go/$version"
        mv "$stage/go" "$BASE/go/$version"
    fi
    ln -sfn "$BASE/go/$version" "$BASE/go/current"
    "$BASE/go/current/bin/go" version
}

install_rust() {
    note 'Install checksum-verified rustup and stable Rust in a dedicated toolchain directory'
    local url digest
    if [[ ! -x $CARGO_HOME/bin/rustup ]]; then
        url="https://static.rust-lang.org/rustup/dist/$RUST_TARGET/rustup-init"
        fetch "$url" "$RUN/rustup-init"
        fetch "$url.sha256" "$RUN/rustup-init.sha256"
        read -r digest _ < "$RUN/rustup-init.sha256"
        [[ $digest =~ ^[a-f0-9]{64}$ ]] || die 'Invalid rustup checksum'
        printf '%s  %s\n' "$digest" "$RUN/rustup-init" | sha256sum --check -
        chmod 700 "$RUN/rustup-init"
        "$RUN/rustup-init" -y --no-modify-path --profile minimal --default-toolchain stable
    fi
    "$CARGO_HOME/bin/rustup" update stable
    "$CARGO_HOME/bin/rustup" default stable
    "$CARGO_HOME/bin/rustc" --version
}

install_tools() {
    note 'Install ProjectDiscovery tool manager and its catalog'
    retry 'Install PDTM' "$BASE/go/current/bin/go" install \
        github.com/projectdiscovery/pdtm/cmd/pdtm@latest || { FAILURES+=(pdtm); return; }
    # Use pdtm's native release handling; no scanning or tool execution against targets.
    retry 'Install all ProjectDiscovery tools' "$BIN/pdtm" -bp "$BASE/pdtm" -ia || \
        FAILURES+=(pdtm-install)
    retry 'Update all ProjectDiscovery tools' "$BIN/pdtm" -bp "$BASE/pdtm" -ua || \
        FAILURES+=(pdtm-update)
    local name
    for name in nuclei httpx subfinder naabu dnsx katana; do
        [[ -x $BASE/pdtm/$name ]] || FAILURES+=("pdtm-missing:$name")
    done
    note 'Install isolated Python CLI applications'
    local apps=(azure-cli shodan pipenv poetry donpapi huggingface-hub)
    for name in "${apps[@]}"; do
        retry "Install Python CLI $name" "$BIN/uv" tool install --upgrade \
            --python 3.12 --managed-python --default-index https://pypi.org/simple "$name" || \
            FAILURES+=("python:$name")
    done
    note 'Install history-derived Rust command-line tools'
    for name in rusthound-ce websocat; do
        retry "Install Rust CLI $name" "$CARGO_HOME/bin/cargo" install --locked "$name" || \
            FAILURES+=("cargo:$name")
    done
    for name in wscat wrangler; do
        retry "Install npm CLI $name" npm install --global --prefix "$BASE/npm" \
            --registry https://registry.npmjs.org "$name@latest" || FAILURES+=("npm:$name")
    done
}

write_environment() {
    note 'Write an opt-in shell environment (existing shell files are untouched)'
    {
        printf '# Generated by cyber-workstation bootstrap. Source from bash or zsh.\n'
        printf 'export CARGO_HOME=%q\n' "$CARGO_HOME"
        printf 'export RUSTUP_HOME=%q\n' "$RUSTUP_HOME"
        printf 'export GOPATH=%q\nexport GOBIN=%q\n' "$GOPATH" "$GOBIN"
        printf 'export UV_PYTHON_INSTALL_DIR=%q\n' "$UV_PYTHON_INSTALL_DIR"
        printf 'export UV_TOOL_DIR=%q\nexport UV_TOOL_BIN_DIR=%q\n' "$UV_TOOL_DIR" "$UV_TOOL_BIN_DIR"
        printf 'export UV_PYTHON_BIN_DIR=%q\n' "$UV_PYTHON_BIN_DIR"
        printf 'export PATH=%q' "$BIN:$BASE/go/current/bin:$CARGO_HOME/bin:$BASE/pdtm:$BASE/npm/bin"
        # shellcheck disable=SC2016
        printf '%s\n' ':"$PATH"'
    } > "$BASE/env.sh.new"
    if [[ -f $BASE/env.sh ]]; then cp -p "$BASE/env.sh" "$RUN/env.sh.before"; fi
    mv "$BASE/env.sh.new" "$BASE/env.sh"
    local rc_file source_line='source /root/.local/share/cyber-workstation/env.sh'
    for rc_file in /root/.zshrc /root/.bashrc; do
        touch "$rc_file"
        grep -Fqx "$source_line" "$rc_file" || printf '\n%s\n' "$source_line" >> "$rc_file"
    done
}

verify_installation() {
    note 'Verify installed runtimes and frequently used commands'
    local command_name
    local commands=(python3 python-latest go rustc cargo docker nmap hashcat pdtm
        nuclei httpx subfinder naabu dnsx katana ffuf feroxbuster sqlmap
        nxc hydra john msfconsole searchsploit az shodan hf donpapi
        rusthound-ce websocat wscat jq tmux)
    : > "$RUN/command-checks.tsv"
    hash -r
    for command_name in "${commands[@]}"; do
        if command -v "$command_name" >/dev/null 2>&1; then
            printf '%s\t%s\n' "$command_name" "$(command -v "$command_name")" >> "$RUN/command-checks.tsv"
        else
            printf '%s\tMISSING\n' "$command_name" >> "$RUN/command-checks.tsv"
            FAILURES+=("command:$command_name")
        fi
    done
    docker version --format '{{.Client.Version}}' >/dev/null 2>&1 || FAILURES+=(docker-client-check)
}

finish() {
    note 'Installation summary'
    dpkg-query -W -f='${binary:Package}\t${Version}\n' > "$RUN/packages-after.tsv"
    "$BIN/uv" tool list > "$RUN/python-tools.txt" 2>&1
    { "$BIN/python-latest" --version; "$BASE/go/current/bin/go" version;
      "$CARGO_HOME/bin/rustc" --version; docker --version; node --version; } > "$RUN/versions.txt" 2>&1
    if ((${#MISSING[@]})); then printf 'Unavailable: %s\n' "${MISSING[@]}"; fi
    if ((${#FAILURES[@]})); then printf 'Failed: %s\n' "${FAILURES[@]}"; fi
    printf '\nEnvironment activation was added to /root/.zshrc and /root/.bashrc.\nLogs: %s\n' "$RUN"
    printf 'Latest Python: python-latest (system python3 remains distribution-managed).\n'
    printf 'Docker is configured for root use.\n'
    printf '\nRECOMMENDATION: Reboot the system before beginning work.\n'
    if ((${#MISSING[@]} + ${#FAILURES[@]})); then return 2; fi
}

main() {
    parse_args "$@"
    detect_platform
    package_plan
    note "Installing complete workstation on ${PRETTY_NAME:-$ID} / $ARCH"
    printf 'This includes large Kali tool and wordlist collections.\n'
    mkdir -p "$BASE" "$BIN" "$BASE/pdtm" "$BASE/runs"
    exec 9> "$BASE/bootstrap.lock"
    flock -n 9 || die 'Another bootstrap for this account is running'
    RUN=$(mktemp -d "$BASE/runs/run-$(date -u +%Y%m%dT%H%M%SZ).XXXXXX")
    exec > >(tee -a "$RUN/install.log") 2>&1
    trap on_error ERR
    note 'Installation begins'
    configure_kali_repository
    install_packages
    install_command_aliases
    retry 'Install current Python runtimes' install_python || FAILURES+=(python-runtime)
    retry 'Install current Go runtime' install_go || FAILURES+=(go-runtime)
    retry 'Install stable Rust runtime' install_rust || FAILURES+=(rust-runtime)
    if [[ -x $BASE/go/current/bin/go && -x $BIN/uv && -x $CARGO_HOME/bin/cargo ]]; then
        install_tools
    else
        FAILURES+=(tool-installers-skipped)
    fi
    write_environment
    if [[ -d /run/systemd/system ]] && command -v systemctl >/dev/null; then
        retry 'Enable and start Docker' systemctl enable --now docker || FAILURES+=(docker-service)
    else
        MISSING+=(docker-service-no-systemd)
    fi
    verify_installation
    # A partial result is intentional, not an unhandled shell failure.
    local result=0
    finish || result=$?
    return "$result"
}

if [[ ${BASH_SOURCE[0]} == "$0" ]]; then main "$@"; fi
