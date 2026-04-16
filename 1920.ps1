#Requires -Version 5.1
[CmdletBinding()]
param(
    [uint32]$DisplayId,
    [switch]$ListDisplays
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$targetWidth = 1920
$targetHeight = 1080
$preferredRefreshRates = @(60.0, 59.99, 59.94, 59.0)
$minimumModuleVersion = [Version]'5.2.0'

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[ OK ] $Message" -ForegroundColor Green
}

function Write-WarnLine {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory)]
        [psobject]$InputObject,
        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $null
}

function Get-DisplayLabel {
    param([Parameter(Mandatory)][psobject]$Display)

    $name = Get-ObjectPropertyValue -InputObject $Display -Names @('DisplayName', 'FriendlyName', 'MonitorName', 'DeviceName')
    if ([string]::IsNullOrWhiteSpace([string]$name)) {
        return "Display $($Display.DisplayId)"
    }

    return [string]$name
}

function Get-CurrentDisplayState {
    param([Parameter(Mandatory)][uint32]$TargetDisplayId)

    $display = Get-DisplayInfo | Where-Object { $_.DisplayId -eq $TargetDisplayId } | Select-Object -First 1
    if ($null -eq $display) {
        throw "DisplayId $TargetDisplayId was not found while reading the current display state."
    }

    $modeString = [string](Get-ObjectPropertyValue -InputObject $display -Names @('Mode'))
    $width = $null
    $height = $null
    $refresh = [double](Get-ObjectPropertyValue -InputObject $display -Names @('RefreshRate', 'Frequency'))

    if ($modeString -match '(?<width>\d+)x(?<height>\d+)@(?<refresh>[\d\.,]+)\s*Hz') {
        $width = [int]$matches.width
        $height = [int]$matches.height
        $refresh = [double]($matches.refresh -replace ',', '.')
    }

    [pscustomobject]@{
        Display   = $display
        Width     = $width
        Height    = $height
        RefreshHz = $refresh
        Mode      = $modeString
    }
}

function Ensure-WindowsDesktop {
    $isWindowsPlatform = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
    if (-not $isWindowsPlatform) {
        throw 'This script only supports Windows.'
    }

    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    if ($os.ProductType -ne 1) {
        Write-WarnLine "This script was designed for Windows desktop editions. Detected: $($os.Caption)"
    }
}

function Ensure-DisplayConfigModule {
    $availableModule = Get-Module -ListAvailable -Name DisplayConfig |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $availableModule -or $availableModule.Version -lt $minimumModuleVersion) {
        Write-Info "Installing DisplayConfig $minimumModuleVersion or newer for the current user..."

        try {
            if (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue) {
                Install-PSResource -Name DisplayConfig -MinimumVersion $minimumModuleVersion -Scope CurrentUser -TrustRepository -Quiet
            }
            else {
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
                }

                $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
                if ($null -ne $psGallery -and $psGallery.InstallationPolicy -ne 'Trusted') {
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                }

                Install-Module -Name DisplayConfig -MinimumVersion $minimumModuleVersion -Scope CurrentUser -Force -AllowClobber
            }
        }
        catch {
            throw "Unable to install the DisplayConfig module automatically. Open PowerShell as your user and run either 'Install-PSResource -Name DisplayConfig -MinimumVersion $minimumModuleVersion -Scope CurrentUser -TrustRepository' or 'Install-Module -Name DisplayConfig -MinimumVersion $minimumModuleVersion -Scope CurrentUser'. Original error: $($_.Exception.Message)"
        }
    }

    Import-Module DisplayConfig -MinimumVersion $minimumModuleVersion -ErrorAction Stop
}

function Show-Displays {
    $displays = Get-DisplayInfo | Sort-Object DisplayId
    if (-not $displays) {
        throw 'No displays were detected by DisplayConfig.'
    }

    $displayRows = foreach ($display in $displays) {
        [pscustomobject]@{
            DisplayId = Get-ObjectPropertyValue -InputObject $display -Names @('DisplayId')
            Name      = Get-DisplayLabel -Display $display
            Primary   = Get-ObjectPropertyValue -InputObject $display -Names @('Primary')
            Active    = Get-ObjectPropertyValue -InputObject $display -Names @('Active', 'IsActive')
            Width     = Get-ObjectPropertyValue -InputObject $display -Names @('Width', 'ResolutionWidth')
            Height    = Get-ObjectPropertyValue -InputObject $display -Names @('Height', 'ResolutionHeight')
            RefreshHz = Get-ObjectPropertyValue -InputObject $display -Names @('RefreshRate', 'Frequency')
            Output    = Get-ObjectPropertyValue -InputObject $display -Names @('ConnectionType', 'OutputTechnology')
        }
    }

    $displayRows | Format-Table -AutoSize | Out-Host
}

function Select-TargetDisplay {
    param([uint32]$RequestedDisplayId)

    $displays = @(Get-DisplayInfo)
    if (-not $displays) {
        throw 'No displays were detected by DisplayConfig.'
    }

    if ($PSBoundParameters.ContainsKey('RequestedDisplayId')) {
        $selected = $displays | Where-Object { $_.DisplayId -eq $RequestedDisplayId } | Select-Object -First 1
        if ($null -eq $selected) {
            Show-Displays
            throw "DisplayId $RequestedDisplayId was not found."
        }

        return $selected
    }

    $selected =
        $displays | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -Names @('Primary')) -eq $true } | Select-Object -First 1

    if ($null -eq $selected) {
        $selected =
            $displays | Where-Object { (Get-ObjectPropertyValue -InputObject $_ -Names @('Active', 'IsActive')) -eq $true } | Select-Object -First 1
    }

    if ($null -eq $selected) {
        $selected = $displays | Select-Object -First 1
    }

    return $selected
}

function Set-DisplayModeWithRefreshFallback {
    param(
        [Parameter(Mandatory)]
        [uint32]$TargetDisplayId,
        [Parameter(Mandatory)]
        [int]$Width,
        [Parameter(Mandatory)]
        [int]$Height,
        [Parameter(Mandatory)]
        [double[]]$RefreshRates,
        [switch]$ChangeAspectRatio
    )

    $errors = @()

    foreach ($refreshRate in $RefreshRates) {
        try {
            Write-Info "Trying refresh rate $refreshRate Hz on ${Width}x${Height}..."

            # For this downscaled mode the DisplayConfig pipeline can keep the old refresh rate.
            # Apply the refresh directly, then verify against the reported Mode string.
            Set-DisplayRefreshRate -DisplayId $TargetDisplayId -RefreshRate $refreshRate -ErrorAction Stop | Out-Null

            $currentState = Get-CurrentDisplayState -TargetDisplayId $TargetDisplayId
            if ($currentState.Width -ne $Width -or $currentState.Height -ne $Height) {
                throw "Refresh change altered the resolution unexpectedly. Current mode: '$($currentState.Mode)'."
            }

            if ([math]::Abs($currentState.RefreshHz - $refreshRate) -lt 0.05) {
                return $currentState.RefreshHz
            }

            throw "Display reported '$($currentState.Mode)' after requesting $refreshRate Hz."
        }
        catch {
            $errors += "$refreshRate Hz: $($_.Exception.Message)"
        }
    }

    throw "None of the candidate refresh rates worked. Tried: $($RefreshRates -join ', ') Hz. Errors: $($errors -join ' | ')"
}

Ensure-WindowsDesktop
Ensure-DisplayConfigModule

if ($ListDisplays) {
    Show-Displays
    return
}

$selectedDisplay = Select-TargetDisplay @PSBoundParameters
$displayLabel = Get-DisplayLabel -Display $selectedDisplay

Write-Info "Target display: $displayLabel (DisplayId $($selectedDisplay.DisplayId))"
Write-Info "Requested mode: ${targetWidth}x${targetHeight} with aspect-ratio correction"
Write-Info "Refresh preference order: $($preferredRefreshRates -join ', ') Hz"

try {
    Write-Info "Applying ${targetWidth}x${targetHeight} first..."
    Get-DisplayConfig |
        Set-DisplayResolution -DisplayId $selectedDisplay.DisplayId -Width $targetWidth -Height $targetHeight -ChangeAspectRatio |
        Use-DisplayConfig -AllowChanges -ErrorAction Stop

    $stateAfterResolution = Get-CurrentDisplayState -TargetDisplayId $selectedDisplay.DisplayId
    Write-Info "Current mode after resolution change: $($stateAfterResolution.Mode)"

    if ($stateAfterResolution.Width -ne $targetWidth -or $stateAfterResolution.Height -ne $targetHeight) {
        throw "Resolution verification failed. Current mode is '$($stateAfterResolution.Mode)'."
    }

    $appliedRefreshRate = Set-DisplayModeWithRefreshFallback -TargetDisplayId $selectedDisplay.DisplayId -Width $targetWidth -Height $targetHeight -RefreshRates $preferredRefreshRates -ChangeAspectRatio

    $finalState = Get-CurrentDisplayState -TargetDisplayId $selectedDisplay.DisplayId
    if ($finalState.Width -ne $targetWidth -or $finalState.Height -ne $targetHeight) {
        throw "Final verification failed because the resolution changed unexpectedly to '$($finalState.Mode)'."
    }

    $refreshMatched = $false
    foreach ($candidateRate in $preferredRefreshRates) {
        if ([math]::Abs($finalState.RefreshHz - $candidateRate) -lt 0.05) {
            $refreshMatched = $true
            break
        }
    }

    if (-not $refreshMatched) {
        throw "Final verification failed because the refresh rate is '$($finalState.RefreshHz)' Hz instead of a value near 60 Hz. Current mode: '$($finalState.Mode)'."
    }

    Write-Ok "Applied $($finalState.Mode) to $displayLabel."
}
catch {
    throw "Display switch failed. $($_.Exception.Message)"
}
