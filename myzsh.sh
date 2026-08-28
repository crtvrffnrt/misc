#!/usr/bin/env bash
# myzsh.sh - Reproduce a polished, root-friendly Zsh setup on Debian systems.
# Safe to run repeatedly. Existing unmanaged dotfiles are backed up before replacement.

set -Eeuo pipefail
IFS=$'\n\t'

readonly MANAGED_MARKER='# Managed by myzsh.sh'
INSTALL_FOR_ALL=false
CHANGE_SHELL=true
UPDATE_REPOS=true
APT_UPDATED=false
LOCK_DIR=''
BACKUP_STAMP=$(date -u +%Y%m%dT%H%M%SZ)

log()  { printf '[myzsh] %s\n' "$*" >&2; }
warn() { printf '[myzsh] WARNING: %s\n' "$*" >&2; }
die()  { printf '[myzsh] ERROR: %s\n' "$*" >&2; exit 1; }

on_error() {
  local rc=$?
  printf '[myzsh] ERROR: command failed at line %s (exit %s): %s\n' \
    "${BASH_LINENO[0]:-?}" "$rc" "${BASH_COMMAND:-unknown}" >&2
  exit "$rc"
}
cleanup() {
  if [[ -n ${LOCK_DIR:-} && -d $LOCK_DIR ]]; then
    rmdir "$LOCK_DIR" 2>/dev/null || true
  fi
}
trap on_error ERR
trap cleanup EXIT INT TERM

usage() {
  cat <<'USAGE'
Usage: myzsh.sh [OPTIONS]

Install a portable Oh My Zsh + Powerlevel10k environment for root/current user.

Options:
  --install-for-all-user  Also configure human users whose login shell is bash/zsh.
  --no-change-shell       Do not change users' login shell to zsh.
  --no-update             Do not update existing clean Git checkouts.
  -h, --help              Show this help.

This script supports Debian-family systems using apt-get and must run as root.
It never imports existing history or personal alias files.
USAGE
}

while (($#)); do
  case $1 in
    --install-for-all-user|--install-for-all-users) INSTALL_FOR_ALL=true ;;
    --no-change-shell) CHANGE_SHELL=false ;;
    --no-update) UPDATE_REPOS=false ;;
    -h|--help) usage; exit 0 ;;
    *) die "unknown option: $1 (try --help)" ;;
  esac
  shift
done

[[ $(id -u) -eq 0 ]] || die 'run this installer as root (for example: sudo bash myzsh.sh)'
[[ -r /etc/os-release ]] || die 'cannot identify this operating system'
# shellcheck disable=SC1091
. /etc/os-release
case ${ID:-}:${ID_LIKE:-} in
  debian:*|ubuntu:*|*:debian*|*:ubuntu*) ;;
  *) die "unsupported system '${PRETTY_NAME:-unknown}'; a Debian-family OS is required" ;;
esac
command -v apt-get >/dev/null 2>&1 || die 'apt-get was not found'
command -v getent >/dev/null 2>&1 || die 'getent was not found'

LOCK_DIR=/run/lock/myzsh-installer.lock
mkdir -p /run/lock
if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  die "another installation appears to be running ($LOCK_DIR exists)"
fi

apt_update_once() {
  if ! $APT_UPDATED; then
    log 'Refreshing APT package metadata...'
    DEBIAN_FRONTEND=noninteractive apt-get update
    APT_UPDATED=true
  fi
}

apt_install_required() {
  local missing=() pkg
  for pkg in "$@"; do
    dpkg-query -W -f='${db:Status-Abbrev}' "$pkg" 2>/dev/null | grep -q '^ii ' || missing+=("$pkg")
  done
  ((${#missing[@]})) || return 0
  apt_update_once
  log "Installing required packages: ${missing[*]}"
  DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends "${missing[@]}"
}

apt_install_if_available() {
  local pkg=$1
  if dpkg-query -W -f='${db:Status-Abbrev}' "$pkg" 2>/dev/null | grep -q '^ii '; then
    return 0
  fi
  apt_update_once
  if apt-cache show "$pkg" >/dev/null 2>&1; then
    log "Installing optional package: $pkg"
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends "$pkg" || \
      warn "optional package '$pkg' could not be installed"
  else
    return 1
  fi
}

apt_install_required zsh git curl ca-certificates fzf
# eza is available on newer Debian releases; exa is the compatible fallback on older ones.
if ! command -v eza >/dev/null 2>&1 && ! command -v exa >/dev/null 2>&1; then
  apt_install_if_available eza || apt_install_if_available exa || \
    warn 'neither eza nor exa is packaged here; FZF previews will use ls'
fi
apt_install_if_available locales || true
apt_install_if_available fonts-powerline || true

ZSH_BIN=$(command -v zsh)
[[ -x $ZSH_BIN ]] || die 'zsh installation verification failed'
grep -Fxq "$ZSH_BIN" /etc/shells || printf '%s\n' "$ZSH_BIN" >>/etc/shells

clone_or_update() {
  local repo=$1 dest=$2 owner=$3
  if [[ -d $dest/.git ]]; then
    local actual
    # The installer runs as root while user-owned repositories must remain user-owned.
    # Scope Git's safe.directory exception to this exact checkout and this invocation.
    actual=$(git -c safe.directory="$dest" -C "$dest" remote get-url origin 2>/dev/null || true)
    if [[ $actual != "$repo" && ${actual%.git} != "${repo%.git}" ]]; then
      warn "$dest has a different origin ($actual); preserving it"
      return 0
    fi
    if $UPDATE_REPOS; then
      if [[ -z $(git -c safe.directory="$dest" -C "$dest" status --porcelain --untracked-files=no 2>/dev/null) ]]; then
        log "Updating $dest"
        git -c safe.directory="$dest" -C "$dest" pull --ff-only --quiet || \
          warn "could not fast-forward $dest; keeping current checkout"
      else
        warn "$dest has local changes; not updating it"
      fi
    fi
  elif [[ -e $dest ]]; then
    local moved="${dest}.pre-myzsh-${BACKUP_STAMP}"
    warn "$dest is not the expected Git checkout; moving it to $moved"
    mv -- "$dest" "$moved"
    git clone --depth=1 --quiet "$repo" "$dest"
  else
    mkdir -p "$(dirname "$dest")"
    git clone --depth=1 --quiet "$repo" "$dest"
  fi
  chown -R "$owner" "$dest"
}

backup_unmanaged() {
  local file=$1 owner=$2
  [[ -e $file || -L $file ]] || return 0
  if [[ -f $file ]] && grep -Fqx "$MANAGED_MARKER" "$file" 2>/dev/null; then
    return 0
  fi
  local backup="${file}.pre-myzsh-${BACKUP_STAMP}"
  cp -a -- "$file" "$backup"
  chown -h "$owner" "$backup"
  log "Backed up $file to $backup"
}

atomic_from_stdin() {
  local dest=$1 owner=$2 mode=$3 dir tmp
  dir=$(dirname "$dest")
  mkdir -p "$dir"
  tmp=$(mktemp "${dir}/.myzsh.tmp.XXXXXX")
  cat >"$tmp"
  chmod "$mode" "$tmp"
  chown "$owner" "$tmp"
  mv -f -- "$tmp" "$dest"
}

choose_dir_color() {
  local state_file=$1 owner=$2 color=''
  if [[ -r $state_file ]]; then
    read -r color <"$state_file" || true
  fi
  case $color in
    22|23|28|29|52|53|54|58|88|89|90|94|95|96|130|131|132|133|134|135) ;;
    *)
      local palette=(22 23 28 29 52 53 54 58 88 89 90 94 95 96 130 131 132 133 134 135)
      local number
      number=$(od -An -N4 -tu4 /dev/urandom | tr -d ' ')
      color=${palette[number % ${#palette[@]}]}
      mkdir -p "$(dirname "$state_file")"
      printf '%s\n' "$color" | atomic_from_stdin "$state_file" "$owner" 600
      ;;
  esac
  printf '%s\n' "$color"
}

write_zshrc() {
  local home=$1 owner=$2
  backup_unmanaged "$home/.zshrc" "$owner"
  atomic_from_stdin "$home/.zshrc" "$owner" 644 <<'ZSHRC'
# Managed by myzsh.sh
# Portable interactive configuration: no imported aliases, secrets, or history content.

# Powerlevel10k instant prompt. Keep near the top of this file.
if [[ -r "${XDG_CACHE_HOME:-$HOME/.cache}/p10k-instant-prompt-${(%):-%n}.zsh" ]]; then
  source "${XDG_CACHE_HOME:-$HOME/.cache}/p10k-instant-prompt-${(%):-%n}.zsh"
fi

export ZSH="$HOME/.oh-my-zsh"
ZSH_THEME='powerlevel10k/powerlevel10k'
zstyle ':omz:update' mode reminder
zstyle ':omz:update' frequency 14
DISABLE_AUTO_UPDATE=true
DISABLE_MAGIC_FUNCTIONS=true

# Syntax highlighting must remain the final plugin in this list.
plugins=(
  git
  copyfile
  z
  zsh-edit-select
  fzf-tab
  zsh-autosuggestions
  zsh-history-substring-search
  zsh-syntax-highlighting
)
source "$ZSH/oh-my-zsh.sh"

# Completion: forgiving matching, descriptions and FZF-backed selection.
zstyle ':completion:*' matcher-list 'm:{a-zA-Z}={A-Za-z}' 'r:|[._-]=* r:|=*'
zstyle ':completion:*' menu no
zstyle ':completion:*' group-name ''
zstyle ':completion:*:descriptions' format '[%d]'
zstyle ':fzf-tab:*' fzf-flags '--height=66%' '--preview-window=right:65%:wrap'
zstyle ':fzf-tab:complete:cd:*' fzf-preview \
  'if command -v eza >/dev/null; then eza -la --color=always --group-directories-first -- "$realpath"; elif command -v exa >/dev/null; then exa -la --color=always --group-directories-first -- "$realpath"; else ls -la --color=always -- "$realpath"; fi'
zstyle ':fzf-tab:complete:(ls|eza|exa|_files):*' fzf-preview \
  'if command -v eza >/dev/null; then eza -lad --color=always -- "$realpath"; [[ -d "$realpath" ]] && eza -la --color=always --group-directories-first -- "$realpath"; elif command -v exa >/dev/null; then exa -lad --color=always -- "$realpath"; [[ -d "$realpath" ]] && exa -la --color=always --group-directories-first -- "$realpath"; else ls -lad --color=always -- "$realpath"; [[ -d "$realpath" ]] && ls -la --color=always -- "$realpath"; fi'

# A fresh history file is created naturally; existing history is never copied by the installer.
HISTFILE="$HOME/.zsh_history"
HISTSIZE=100000
SAVEHIST=100000
setopt APPEND_HISTORY INC_APPEND_HISTORY SHARE_HISTORY
setopt HIST_EXPIRE_DUPS_FIRST HIST_IGNORE_DUPS HIST_IGNORE_SPACE HIST_VERIFY
setopt GLOB_DOTS AUTO_CD INTERACTIVE_COMMENTS

# Autosuggestion and substring-search appearance/key bindings.
ZSH_AUTOSUGGEST_HIGHLIGHT_STYLE='fg=244'
bindkey '^[[A' history-substring-search-up
bindkey '^[[B' history-substring-search-down
bindkey '^[OA' history-substring-search-up
bindkey '^[OB' history-substring-search-down

# Debian's fzf package ships integration here. Newer upstream fzf supports `fzf --zsh`.
if [[ -r /usr/share/doc/fzf/examples/key-bindings.zsh ]]; then
  source /usr/share/doc/fzf/examples/key-bindings.zsh
elif [[ -r /usr/share/fzf/key-bindings.zsh ]]; then
  source /usr/share/fzf/key-bindings.zsh
elif (( $+commands[fzf] )) && fzf --zsh >/dev/null 2>&1; then
  source <(fzf --zsh)
fi
export FZF_DEFAULT_OPTS="${FZF_DEFAULT_OPTS:+$FZF_DEFAULT_OPTS }--height=66% --layout=reverse --border"
(( $+widgets[fzf-history-widget] )) && bindkey '^R' fzf-history-widget
bindkey -r '^S' 2>/dev/null || true
[[ -t 0 ]] && stty -ixon -ixoff 2>/dev/null || true

# Optional Debian command-not-found integration.
[[ -r /etc/zsh_command_not_found ]] && source /etc/zsh_command_not_found

[[ -r "$HOME/.p10k.zsh" ]] && source "$HOME/.p10k.zsh"
ZSHRC
}

write_p10k() {
  local home=$1 owner=$2 dir_bg=$3
  backup_unmanaged "$home/.p10k.zsh" "$owner"
  atomic_from_stdin "$home/.p10k.zsh" "$owner" 644 <<P10K
# Managed by myzsh.sh
# Compact Powerlevel10k configuration based on this machine's two-line rainbow prompt.
# The directory background ($dir_bg) is randomly selected once and persisted per user.

() {
  emulate -L zsh -o extended_glob
  unset -m '(POWERLEVEL9K_*|DEFAULT_USER)~POWERLEVEL9K_GITSTATUS_DIR'
  [[ \$ZSH_VERSION == (5.<1->*|<6->.*) ]] || return

  typeset -g POWERLEVEL9K_MODE=ascii
  typeset -g POWERLEVEL9K_ICON_PADDING=none
  typeset -g POWERLEVEL9K_LEFT_PROMPT_ELEMENTS=(dir vcs newline prompt_char)
  typeset -g POWERLEVEL9K_RIGHT_PROMPT_ELEMENTS=(
    status command_execution_time background_jobs virtualenv
    kubecontext aws azure context time newline
  )

  typeset -g POWERLEVEL9K_PROMPT_ADD_NEWLINE=true
  typeset -g POWERLEVEL9K_MULTILINE_FIRST_PROMPT_GAP_CHAR='-'
  typeset -g POWERLEVEL9K_MULTILINE_FIRST_PROMPT_GAP_FOREGROUND=240
  typeset -g POWERLEVEL9K_MULTILINE_FIRST_PROMPT_PREFIX=
  typeset -g POWERLEVEL9K_MULTILINE_NEWLINE_PROMPT_PREFIX=
  typeset -g POWERLEVEL9K_MULTILINE_LAST_PROMPT_PREFIX=
  typeset -g POWERLEVEL9K_LEFT_SEGMENT_SEPARATOR=''
  typeset -g POWERLEVEL9K_LEFT_SUBSEGMENT_SEPARATOR=' '
  typeset -g POWERLEVEL9K_RIGHT_SEGMENT_SEPARATOR=''
  typeset -g POWERLEVEL9K_RIGHT_SUBSEGMENT_SEPARATOR=' '

  typeset -g POWERLEVEL9K_PROMPT_CHAR_BACKGROUND=
  typeset -g POWERLEVEL9K_PROMPT_CHAR_OK_VIINS_FOREGROUND=76
  typeset -g POWERLEVEL9K_PROMPT_CHAR_ERROR_VIINS_FOREGROUND=196
  typeset -g POWERLEVEL9K_PROMPT_CHAR_OK_VIINS_CONTENT_EXPANSION='>'
  typeset -g POWERLEVEL9K_PROMPT_CHAR_ERROR_VIINS_CONTENT_EXPANSION='>'

  typeset -g POWERLEVEL9K_DIR_BACKGROUND=$dir_bg
  typeset -g POWERLEVEL9K_DIR_FOREGROUND=255
  typeset -g POWERLEVEL9K_DIR_SHORTENED_FOREGROUND=252
  typeset -g POWERLEVEL9K_DIR_ANCHOR_FOREGROUND=231
  typeset -g POWERLEVEL9K_DIR_ANCHOR_BOLD=true
  typeset -g POWERLEVEL9K_SHORTEN_STRATEGY=truncate_to_unique
  typeset -g POWERLEVEL9K_SHORTEN_DELIMITER=
  typeset -g POWERLEVEL9K_SHORTEN_DIR_LENGTH=1
  typeset -g POWERLEVEL9K_DIR_MAX_LENGTH=80
  typeset -g POWERLEVEL9K_DIR_MIN_COMMAND_COLUMNS=40
  typeset -g POWERLEVEL9K_DIR_MIN_COMMAND_COLUMNS_PCT=50
  typeset -g POWERLEVEL9K_DIR_HYPERLINK=false
  typeset -g POWERLEVEL9K_DIR_SHOW_WRITABLE=v3
  local anchor_files=(.git .hg .svn .terraform .tool-versions .python-version .node-version Cargo.toml go.mod package.json)
  typeset -g POWERLEVEL9K_SHORTEN_FOLDER_MARKER="(\${(j:|:)anchor_files})"

  typeset -g POWERLEVEL9K_VCS_CLEAN_BACKGROUND=2
  typeset -g POWERLEVEL9K_VCS_MODIFIED_BACKGROUND=3
  typeset -g POWERLEVEL9K_VCS_UNTRACKED_BACKGROUND=2
  typeset -g POWERLEVEL9K_VCS_CONFLICTED_BACKGROUND=1
  typeset -g POWERLEVEL9K_VCS_LOADING_BACKGROUND=8
  typeset -g POWERLEVEL9K_VCS_BRANCH_ICON=
  typeset -g POWERLEVEL9K_VCS_UNTRACKED_ICON='?'
  typeset -g POWERLEVEL9K_VCS_BACKENDS=(git)
  typeset -g POWERLEVEL9K_VCS_MAX_INDEX_SIZE_DIRTY=-1
  typeset -g POWERLEVEL9K_VCS_DISABLED_WORKDIR_PATTERN='~'

  typeset -g POWERLEVEL9K_STATUS_EXTENDED_STATES=true
  typeset -g POWERLEVEL9K_STATUS_OK=false
  typeset -g POWERLEVEL9K_STATUS_ERROR=false
  typeset -g POWERLEVEL9K_STATUS_OK_PIPE=true
  typeset -g POWERLEVEL9K_STATUS_ERROR_PIPE=true
  typeset -g POWERLEVEL9K_STATUS_ERROR_SIGNAL=true
  typeset -g POWERLEVEL9K_STATUS_VERBOSE_SIGNAME=false
  typeset -g POWERLEVEL9K_STATUS_OK_PIPE_FOREGROUND=2
  typeset -g POWERLEVEL9K_STATUS_OK_PIPE_BACKGROUND=0
  typeset -g POWERLEVEL9K_STATUS_ERROR_PIPE_FOREGROUND=3
  typeset -g POWERLEVEL9K_STATUS_ERROR_PIPE_BACKGROUND=1
  typeset -g POWERLEVEL9K_STATUS_ERROR_SIGNAL_FOREGROUND=3
  typeset -g POWERLEVEL9K_STATUS_ERROR_SIGNAL_BACKGROUND=1

  typeset -g POWERLEVEL9K_COMMAND_EXECUTION_TIME_THRESHOLD=3
  typeset -g POWERLEVEL9K_COMMAND_EXECUTION_TIME_PRECISION=0
  typeset -g POWERLEVEL9K_COMMAND_EXECUTION_TIME_FORMAT='d h m s'
  typeset -g POWERLEVEL9K_COMMAND_EXECUTION_TIME_FOREGROUND=0
  typeset -g POWERLEVEL9K_COMMAND_EXECUTION_TIME_BACKGROUND=3
  typeset -g POWERLEVEL9K_BACKGROUND_JOBS_FOREGROUND=6
  typeset -g POWERLEVEL9K_BACKGROUND_JOBS_BACKGROUND=0
  typeset -g POWERLEVEL9K_BACKGROUND_JOBS_VERBOSE=false

  typeset -g POWERLEVEL9K_VIRTUALENV_FOREGROUND=0
  typeset -g POWERLEVEL9K_VIRTUALENV_BACKGROUND=4
  typeset -g POWERLEVEL9K_VIRTUALENV_SHOW_PYTHON_VERSION=false
  typeset -g POWERLEVEL9K_KUBECONTEXT_SHOW_ON_COMMAND='kubectl|helm|kubectx|kubens|oc|k9s'
  typeset -g POWERLEVEL9K_KUBECONTEXT_DEFAULT_FOREGROUND=7
  typeset -g POWERLEVEL9K_KUBECONTEXT_DEFAULT_BACKGROUND=5
  typeset -g POWERLEVEL9K_AWS_SHOW_ON_COMMAND='aws|terraform|tofu|pulumi|terragrunt'
  typeset -g POWERLEVEL9K_AWS_DEFAULT_FOREGROUND=7
  typeset -g POWERLEVEL9K_AWS_DEFAULT_BACKGROUND=1
  typeset -g POWERLEVEL9K_AZURE_SHOW_ON_COMMAND='az|terraform|tofu|pulumi|terragrunt'
  typeset -g POWERLEVEL9K_AZURE_OTHER_FOREGROUND=7
  typeset -g POWERLEVEL9K_AZURE_OTHER_BACKGROUND=4

  typeset -g POWERLEVEL9K_CONTEXT_ROOT_FOREGROUND=1
  typeset -g POWERLEVEL9K_CONTEXT_ROOT_BACKGROUND=0
  typeset -g POWERLEVEL9K_CONTEXT_REMOTE_FOREGROUND=3
  typeset -g POWERLEVEL9K_CONTEXT_REMOTE_BACKGROUND=0
  typeset -g POWERLEVEL9K_CONTEXT_ROOT_TEMPLATE='%n@%m'
  typeset -g POWERLEVEL9K_CONTEXT_REMOTE_TEMPLATE='%n@%m'
  typeset -g POWERLEVEL9K_CONTEXT_DEFAULT_CONTENT_EXPANSION=
  typeset -g POWERLEVEL9K_CONTEXT_DEFAULT_VISUAL_IDENTIFIER_EXPANSION=
  typeset -g POWERLEVEL9K_TIME_FOREGROUND=0
  typeset -g POWERLEVEL9K_TIME_BACKGROUND=7
  typeset -g POWERLEVEL9K_TIME_FORMAT='%D{%H:%M:%S}'
  typeset -g POWERLEVEL9K_TIME_UPDATE_ON_COMMAND=false

  typeset -g POWERLEVEL9K_TRANSIENT_PROMPT=off
  typeset -g POWERLEVEL9K_INSTANT_PROMPT=quiet
  typeset -g POWERLEVEL9K_DISABLE_HOT_RELOAD=true
} always
P10K
}

verify_user_install() {
  local user=$1 home=$2
  local required=(
    "$home/.zshrc"
    "$home/.p10k.zsh"
    "$home/.oh-my-zsh/oh-my-zsh.sh"
    "$home/.oh-my-zsh/custom/themes/powerlevel10k/powerlevel10k.zsh-theme"
    "$home/.oh-my-zsh/custom/plugins/fzf-tab/fzf-tab.plugin.zsh"
    "$home/.oh-my-zsh/custom/plugins/zsh-autosuggestions/zsh-autosuggestions.plugin.zsh"
    "$home/.oh-my-zsh/custom/plugins/zsh-history-substring-search/zsh-history-substring-search.plugin.zsh"
    "$home/.oh-my-zsh/custom/plugins/zsh-edit-select/zsh-edit-select.plugin.zsh"
    "$home/.oh-my-zsh/custom/plugins/zsh-syntax-highlighting/zsh-syntax-highlighting.plugin.zsh"
  )
  local item
  for item in "${required[@]}"; do
    [[ -r $item ]] || die "verification failed for $user: missing $item"
  done
  "$ZSH_BIN" -n "$home/.zshrc"
  "$ZSH_BIN" -n "$home/.p10k.zsh"
}

configure_user() {
  local user=$1 home=$2 uid=$3 gid=$4 owner
  owner="${uid}:${gid}"
  [[ -n $home && $home == /* && $home != / ]] || { warn "skipping $user: unsafe home '$home'"; return; }
  if [[ ! -d $home ]]; then
    warn "skipping $user: home directory does not exist ($home)"
    return
  fi
  log "Configuring Zsh for $user ($home)"

  clone_or_update https://github.com/ohmyzsh/ohmyzsh.git "$home/.oh-my-zsh" "$owner"
  local custom="$home/.oh-my-zsh/custom"
  clone_or_update https://github.com/romkatv/powerlevel10k.git "$custom/themes/powerlevel10k" "$owner"
  clone_or_update https://github.com/Aloxaf/fzf-tab.git "$custom/plugins/fzf-tab" "$owner"
  clone_or_update https://github.com/zsh-users/zsh-autosuggestions.git "$custom/plugins/zsh-autosuggestions" "$owner"
  clone_or_update https://github.com/zsh-users/zsh-history-substring-search.git "$custom/plugins/zsh-history-substring-search" "$owner"
  clone_or_update https://github.com/Michael-Matta1/zsh-edit-select.git "$custom/plugins/zsh-edit-select" "$owner"
  clone_or_update https://github.com/zsh-users/zsh-syntax-highlighting.git "$custom/plugins/zsh-syntax-highlighting" "$owner"

  local color
  color=$(choose_dir_color "$home/.config/myzsh/dir-color" "$owner")
  write_zshrc "$home" "$owner"
  write_p10k "$home" "$owner" "$color"
  verify_user_install "$user" "$home"

  if $CHANGE_SHELL; then
    local current_shell
    current_shell=$(getent passwd "$user" | cut -d: -f7)
    if [[ $current_shell != "$ZSH_BIN" ]]; then
      usermod -s "$ZSH_BIN" "$user"
      [[ $(getent passwd "$user" | cut -d: -f7) == "$ZSH_BIN" ]] || \
        die "failed to change login shell for $user"
    fi
  fi
  log "Verified $user (directory color: $color)"
}

collect_users() {
  local uid_min
  uid_min=$(awk '$1 == "UID_MIN" {print $2; exit}' /etc/login.defs 2>/dev/null || true)
  [[ $uid_min =~ ^[0-9]+$ ]] || uid_min=1000
  printf 'root\n'
  if $INSTALL_FOR_ALL; then
    getent passwd | awk -F: -v min="$uid_min" \
      '($3 >= min) && ($7 ~ /\/(ba|z)sh$/) && ($7 !~ /(nologin|false)$/) {print $1}'
  fi
}

mapfile -t TARGET_USERS < <(collect_users | awk '!seen[$0]++')
((${#TARGET_USERS[@]})) || die 'no eligible users found'

for target_user in "${TARGET_USERS[@]}"; do
  passwd_row=$(getent passwd "$target_user") || { warn "skipping unknown user $target_user"; continue; }
  IFS=: read -r _ _ target_uid target_gid _ target_home _ <<<"$passwd_row"
  configure_user "$target_user" "$target_home" "$target_uid" "$target_gid"
done

log 'All requested users passed configuration and syntax checks.'
log "Start a new login shell with: exec $ZSH_BIN -l"
