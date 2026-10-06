<#
.SYNOPSIS
    Updates installed packages using Winget and Chocolatey with improved logging and exclusions.

.DESCRIPTION
    This script checks for available upgrades using both Winget and Chocolatey (if installed).
    It filters them based on a JSONC/JSON exclusion list and performs upgrades.
    It produces a unified log file for each run and automatically rotates old logs.

.PARAMETER LogPath
    Optional path to a specific log file or directory. If a directory is provided, a timestamped log is created.

.PARAMETER ExclusionsFile
    Path to the exclusions file (.jsonc or .json) containing package IDs to exclude.
    JSONC files support // line comments and /* */ block comments.

.PARAMETER Force
    If specified, skips the confirmation prompt.

.PARAMETER DryRun
    Lists planned upgrades without actually performing them. No packages are modified.

.PARAMETER KeepLogs
    Number of days to retain log files. Logs older than this are deleted. Default: 30.

.PARAMETER SkipWinget
    Skip Winget upgrade checks entirely.

.PARAMETER SkipChocolatey
    Skip Chocolatey upgrade checks entirely.

.EXAMPLE
    .\Update-InstalledPackages.ps1 -Force

.EXAMPLE
    .\Update-InstalledPackages.ps1 -DryRun

.EXAMPLE
    .\Update-InstalledPackages.ps1 -SkipChocolatey -Force -KeepLogs 14
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath,

    [Parameter(Mandatory=$false)]
    [string]$ExclusionsFile,

    [switch]$Force,
    [switch]$DryRun,

    [Parameter(Mandatory=$false)]
    [int]$KeepLogs = 30,

    [switch]$SkipWinget,
    [switch]$SkipChocolatey
)

Add-Type -AssemblyName System.Windows.Forms

# ------------------------
# Configuration & Setup
# ------------------------

# Determine Script Directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) {
    $ScriptDir = $PWD.Path
}

# Determine Log File Path
if (-not $LogPath) {
    $LogBaseDir = Join-Path $ScriptDir "logs"
    if (-not (Test-Path -Path $LogBaseDir)) {
        New-Item -Path $LogBaseDir -ItemType Directory -Force | Out-Null
    }
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:CurrentLogFile = Join-Path $LogBaseDir ("system-upgrade-" + $TimeStamp + ".log")
} elseif (Test-Path -Path $LogPath -PathType Container) {
    $LogBaseDir = $LogPath
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:CurrentLogFile = Join-Path $LogPath ("system-upgrade-" + $TimeStamp + ".log")
} else {
    $LogBaseDir = $null
    $parent = Split-Path -Parent $LogPath
    if (-not (Test-Path -Path $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }
    $Script:CurrentLogFile = $LogPath
}

# Log rotation: remove logs older than $KeepLogs days
if ($LogBaseDir -and (Test-Path $LogBaseDir) -and $KeepLogs -gt 0) {
    $cutoff = (Get-Date).AddDays(-$KeepLogs)
    Get-ChildItem -Path $LogBaseDir -Filter '*.log' |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        }
}

# Determine Exclusions File (.jsonc preferred, falls back to .json)
if (-not $ExclusionsFile) {
    $jsonc = Join-Path $ScriptDir "winget-upgrade-exclusions.jsonc"
    $json  = Join-Path $ScriptDir "winget-upgrade-exclusions.json"
    if (Test-Path $jsonc) {
        $ExclusionsFile = $jsonc
    } elseif (Test-Path $json) {
        $ExclusionsFile = $json
    }
}

# Winget flags as a proper array (avoids fragile string-split)
$WingetAcceptFlags = @('--accept-package-agreements', '--accept-source-agreements')

# Network-related exit codes that are worth retrying
$Script:RetryExitCodes = @(
    -2147012894,  # 0x80072EE2  WINHTTP_ERROR_TIMEOUT
    -2147012867,  # 0x80072EFD  WINHTTP_ERROR_CONNECTION_ERROR
    -2147012866,  # 0x80072EFE  WINHTTP_ERROR_CONNECTION_ABORTED
    -2147023293   # 0x80070643  General installer failure (sometimes transient)
)

# ------------------------
# Functions
# ------------------------

function Assert-AdminPrivilege {
    param($ScriptParameters)
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Warning "Not running as Administrator. Relaunching elevated..."
        $psi = New-Object System.Diagnostics.ProcessStartInfo

        # Prefer pwsh if available, else powershell
        if (Get-Command pwsh -ErrorAction SilentlyContinue) {
            $psi.FileName = 'pwsh.exe'
        } else {
            $psi.FileName = 'powershell.exe'
        }

        # Reconstruct arguments
        $argsList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
        if ($ScriptParameters) {
            $ScriptParameters.GetEnumerator() | ForEach-Object {
                if ($_.Value -is [switch] -and $_.Value) {
                    $argsList += "-$($_.Key)"
                } elseif ($_.Value -isnot [switch]) {
                    $argsList += "-$($_.Key)"
                    $argsList += "`"$($_.Value)`""
                }
            }
        }

        $psi.Arguments = $argsList -join " "
        $psi.Verb = 'runas'

        try {
            [System.Diagnostics.Process]::Start($psi) | Out-Null
        } catch {
            Write-Error "Failed to relaunch elevated: $_"
        }
        Exit
    }
}

function Write-UpdaterLog {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Color = 'Gray'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] $Message"
    Write-Host $line -ForegroundColor $Color
    try {
        Add-Content -Path $Script:CurrentLogFile -Value $line -ErrorAction Stop
    } catch {
        # Fallback if log file isn't writable
        Write-Host " [Error writing to log: $_]" -ForegroundColor Red
    }
}

function Show-GuiConfirmation {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Proceed to upgrade the listed packages?",
        "System Package Upgrade Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

# ------------------------
# Retry Helper
# ------------------------

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Runs a script block and retries on transient network exit codes.
    .PARAMETER ScriptBlock
        The block to execute. Must return an integer exit code via 'return $LASTEXITCODE'.
    .PARAMETER MaxRetries
        Maximum number of retry attempts (default: 3).
    .PARAMETER DelaySeconds
        Seconds to wait between retries (default: 15).
    .PARAMETER PackageName
        Name used in log messages.
    #>
    param(
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,
        [int]$MaxRetries   = 3,
        [int]$DelaySeconds = 15,
        [string]$PackageName = 'package'
    )

    $attempt = 0
    do {
        $attempt++
        $exitCode = & $ScriptBlock

        if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
            return $exitCode   # success
        }

        if ($attempt -le $MaxRetries -and $Script:RetryExitCodes -contains $exitCode) {
            $hex = '0x{0:X8}' -f ([int64]$exitCode -band 0xFFFFFFFF)
            Write-UpdaterLog -Message ("  [Retry {0}/{1}] Transient error ({2}) upgrading '{3}'. Waiting {4}s..." -f $attempt, $MaxRetries, $hex, $PackageName, $DelaySeconds) -Color 'Yellow'
            Start-Sleep -Seconds $DelaySeconds
        } else {
            return $exitCode   # non-retryable or exhausted retries
        }
    } while ($attempt -le $MaxRetries)

    return $exitCode
}

# ------------------------
# Winget Functions
# ------------------------

function Get-WingetUpgrade {
    <#
    .SYNOPSIS
        Wraps the logic of trying JSON first, then failing back to Table parsing.
    #>

    # Check winget availability
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-UpdaterLog -Message "Winget not found. Skipping." -Color "Gray"
        return @()
    }

    Write-UpdaterLog -Message "Querying Winget for available upgrades..." -Color "Cyan"

    $upgrades = @()
    $triedJson = $false

    try {
        $rawLines = winget upgrade --output json 2>&1
        $rawText = $rawLines -join "`n"

        # Heuristic check for JSON support
        if ($rawText -match 'Unknown option|unrecognized|not a valid option|--output' -and -not ($rawText.TrimStart().StartsWith('[') -or $rawText.TrimStart().StartsWith('{'))) {
            throw "winget does not appear to support --output json."
        }

        $triedJson = $true
        # Log raw JSON attempt snippet
        Add-Content -Path $Script:CurrentLogFile -Value "`n===== RAW winget JSON ATTEMPT =====`n"

        $upgrades = ConvertFrom-WingetJson -RawText $rawText

        if ($upgrades) {
            Write-UpdaterLog -Message "Successfully parsed upgrades from winget JSON output." -Color "DarkGray"
        } else {
            throw "Failed to parse JSON from winget output."
        }
    } catch {
        if ($triedJson) {
            Write-UpdaterLog -Message ("Winget JSON attempt failed: {0}. Falling back to table parsing." -f $_.Exception.Message) -Color "DarkGray"
        }

        # Fallback to table parsing
        Write-UpdaterLog -Message "Parsing winget table output..." -Color "DarkGray"
        $raw = winget upgrade 2>&1
        Add-Content -Path $Script:CurrentLogFile -Value "`n===== RAW winget table output =====`n"
        if ($raw) {
            Add-Content -Path $Script:CurrentLogFile -Value ($raw -join "`n")
        }
        $upgrades = ConvertFrom-WingetTable -RawLines $raw
    }

    # Tag with Manager
    $tagged = @()
    if ($upgrades) {
        foreach ($u in $upgrades) {
            $u | Add-Member -NotePropertyName "Manager" -NotePropertyValue "Winget" -Force
            $tagged += $u
        }
    }
    return $tagged
}

function ConvertFrom-WingetJson {
    param([string]$RawText)
    try {
        # robust JSON extraction (handling potential noise before/after)
        $firstOpen = $RawText.IndexOf('[')
        if ($firstOpen -lt 0) { $firstOpen = $RawText.IndexOf('{') }
        if ($firstOpen -lt 0) { return $null }

        $lastCloseArray = $RawText.LastIndexOf(']')
        $lastCloseObject = $RawText.LastIndexOf('}')
        $IsArray = ($RawText[$firstOpen] -eq '[')
        $lastClose = if ($IsArray) { $lastCloseArray } else { $lastCloseObject }

        if ($lastClose -lt $firstOpen) { return $null }

        $jsonCandidate = $RawText.Substring($firstOpen, ($lastClose - $firstOpen + 1))
        $parsed = $jsonCandidate | ConvertFrom-Json -ErrorAction Stop

        $results = @()
        foreach ($item in $parsed) {
            # Unified property access
            $id = if ($item.PSObject.Properties.Match('Id').Count) { $item.Id } elseif ($item.PSObject.Properties.Match('PackageId').Count) { $item.PackageId } else { $null }
            $name = if ($item.PSObject.Properties.Match('Name').Count) { $item.Name } elseif ($item.PSObject.Properties.Match('PackageName').Count) { $item.PackageName } else { $null }
            $ver = if ($item.PSObject.Properties.Match('Version').Count) { $item.Version } else { $null }
            $avail = if ($item.PSObject.Properties.Match('AvailableVersion').Count) { $item.AvailableVersion } else { $null }
            $source = if ($item.PSObject.Properties.Match('Source').Count) { $item.Source } else { $null }

            if ($id -and $name) {
                $results += [PSCustomObject]@{
                    Name      = $name
                    Id        = $id
                    Version   = $ver
                    Available = $avail
                    Source    = $source
                }
            }
        }
        return $results
    } catch {
        Write-UpdaterLog -Message "Error internal parsing JSON: $_" -Color "Red"
        return $null
    }
}

function ConvertFrom-WingetTable {
    param($RawLines)
    try {
        $lines = $RawLines -split "`r?`n"
        $sepMatch = $lines | Select-String '^-{3,}' | Select-Object -First 1
        if (-not $sepMatch) { return $null }

        $sepIndex = $sepMatch.LineNumber - 1

        $pkgLines = $lines[($sepIndex + 1)..($lines.Count - 1)] | Where-Object {
            $_.Trim() -ne '' -and
            $_ -notmatch '^\d+\s+upgrades available' -and
            $_ -notmatch '^\d+\s+package\(s\)\s+have\s+version\s+numbers'
        }

        $results = @()
        foreach ($line in $pkgLines) {
            # Trim trailing source (winget) and spaces
            $l = $line -replace '\s+winget\s*$', ''
            $l = $l.Trim()

            # Split by whitespace
            $parts = $l -split '\s+'

            # Since Id, Version, Available generally don't contain spaces...
            # The last 3 items in the parts array are Available, Version, Id (in reverse)
            # Everything before them is Name.

            if ($parts.Count -ge 4) {
                # We expect at least chunks for: Name..., Id, Version, Available
                $avail = $parts[-1]
                $ver = $parts[-2]
                $id = $parts[-3]

                # Reconstruct Name
                $nameParts = $parts[0..($parts.Count - 4)]
                $name = ($nameParts -join ' ').Trim()

                # Extra check in case name had `<name> [<id>]` format
                if ($name -match '^(.*)\s\[(.*)\]$') {
                    $name = $matches[1]
                }

                $results += [PSCustomObject]@{
                    Name      = $name
                    Id        = $id
                    Version   = $ver
                    Available = $avail
                    Source    = 'winget'
                }
            } elseif ($parts.Count -eq 3) {
                # It's possible to just have Name Id Version if there's no available version info
                $ver = $parts[-1]
                $id = $parts[-2]
                $name = $parts[0].Trim()
                $results += [PSCustomObject]@{
                    Name      = $name
                    Id        = $id
                    Version   = $ver
                    Available = ''
                    Source    = 'winget'
                }
            }
        }
        return $results
    } catch {
        return $null
    }
}

# ------------------------
# Chocolatey Functions
# ------------------------

function Get-ChocolateyUpgrade {
    Write-UpdaterLog -Message "Querying Chocolatey for available upgrades..." -Color "Cyan"

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-UpdaterLog -Message "Chocolatey not found. Skipping." -Color "Gray"
        return @()
    }

    $upgrades = @()
    try {
        # -r for raw output: name|version|new_version|pinned
        $raw = choco outdated -r --ignore-pinned 2>&1
        foreach ($line in $raw) {
            # Skip empty lines or possible other output if not pipe delimited
            if ($line -match '\|') {
                $parts = $line -split '\|'
                if ($parts.Count -ge 3) {
                    $upgrades += [PSCustomObject]@{
                        Name      = $parts[0]
                        Id        = $parts[0] # Chocolatey uses ID as package name
                        Version   = $parts[1]
                        Available = $parts[2]
                        Source    = 'chocolatey'
                        Manager   = 'Chocolatey'
                    }
                }
            }
        }

        if ($upgrades.Count -gt 0) {
            Write-UpdaterLog -Message ("Found {0} Chocolatey upgrades." -f $upgrades.Count) -Color "DarkGray"
        } else {
            Write-UpdaterLog -Message "No Chocolatey upgrades found." -Color "DarkGray"
        }
    } catch {
        Write-UpdaterLog -Message "Error querying Chocolatey: $_" -Color "Red"
    }
    return $upgrades
}

# ------------------------
# Main Update Logic
# ------------------------

function Invoke-PackageUpdate {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ExclusionsFile,

        [Parameter(Mandatory=$false)]
        [switch]$Force,

        [switch]$DryRun,
        [switch]$SkipWinget,
        [switch]$SkipChocolatey
    )

    # 1. Load Exclusions
    $ExcludeIds = @()
    if ($ExclusionsFile -and (Test-Path $ExclusionsFile)) {
        try {
            # Strip JSONC-style comments before parsing so .jsonc files are supported
            $raw = Get-Content -Path $ExclusionsFile -Raw -ErrorAction Stop
            # Remove block comments  /* ... */
            $raw = [regex]::Replace($raw, '/\*.*?\*/', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            # Remove line comments   // ...
            $raw = [regex]::Replace($raw, '//[^\r\n]*', '')
            $ExcludeIds = $raw | ConvertFrom-Json -ErrorAction Stop
            Write-UpdaterLog -Message "Loaded exclusions from $ExclusionsFile" -Color "DarkGray"
        } catch {
            Write-Warning "Could not read '$ExclusionsFile'. Error: $_"
        }
    }

    if ($DryRun) {
        Write-UpdaterLog -Message "*** DRY RUN MODE - no packages will be upgraded ***" -Color "Yellow"
    }

    Write-UpdaterLog -Message "Starting system upgrade session. Log: $Script:CurrentLogFile" -Color "Cyan"

    # 2. Get Upgrades from all sources
    $allUpgrades = @()

    # Winget
    if (-not $SkipWinget) {
        $wingetUpgrades = Get-WingetUpgrade
        if ($wingetUpgrades) { $allUpgrades += $wingetUpgrades }
    } else {
        Write-UpdaterLog -Message "Winget skipped (-SkipWinget)." -Color "Gray"
    }

    # Chocolatey
    if (-not $SkipChocolatey) {
        $chocoUpgrades = Get-ChocolateyUpgrade
        if ($chocoUpgrades) { $allUpgrades += $chocoUpgrades }
    } else {
        Write-UpdaterLog -Message "Chocolatey skipped (-SkipChocolatey)." -Color "Gray"
    }

    if (-not $allUpgrades -or $allUpgrades.Count -eq 0) {
        Write-UpdaterLog -Message "No upgradable packages detected from any source." -Color "Green"
        return
    }

    Write-UpdaterLog -Message ("Found {0} potential upgrades total." -f $allUpgrades.Count) -Color "Cyan"

    # 3. Apply Exclusions
    $toUpgrade = @()
    foreach ($u in $allUpgrades) {
        # Normalize ID for comparison (remove version tags sometimes in ID field e.g. "Vendor.App [Source]")
        $cleanId = ($u.Id -split '[\s\[]')[0]

        $isExcluded = $false
        foreach ($ex in $ExcludeIds) {
            if ($cleanId -ieq $ex -or $u.Id -ieq $ex) {
                $isExcluded = $true; break
            }
        }

        if ($isExcluded) {
            Write-UpdaterLog -Message ("  Skipping excluded: [{0}] {1}" -f $u.Manager, $u.Name) -Color "DarkGray"
        } else {
            $toUpgrade += $u
        }
    }

    if ($toUpgrade.Count -eq 0) {
        Write-UpdaterLog -Message "All available upgrades are excluded." -Color "Yellow"
        return
    }

    # 4. Confirm
    Write-UpdaterLog -Message "Planned upgrades:" -Color "Cyan"
    $toUpgrade | ForEach-Object {
        Write-UpdaterLog -Message (" - [{0}] {1} ({2} -> {3})" -f $_.Manager, $_.Name, $_.Version, $_.Available)
    }

    if ($DryRun) {
        Write-UpdaterLog -Message "Dry run complete. No changes made." -Color "Yellow"
        return
    }

    if (-not $Force) {
        if (-not (Show-GuiConfirmation)) {
            Write-UpdaterLog -Message "User canceled." -Color "Yellow"
            return
        }
    } else {
        Write-UpdaterLog -Message "Force enabled, proceeding..." -Color "Cyan"
    }

    # 5. Execute
    $results = @()
    foreach ($pkg in $toUpgrade) {
        Write-UpdaterLog -Message "Upgrading $($pkg.Name) [$($pkg.Id)] via $($pkg.Manager)..." -Color "Magenta"

        if ($PSCmdlet.ShouldProcess($pkg.Name, "$($pkg.Manager) Upgrade")) {

            $exitCode = 0
            $status = 'Failed'

            try {
                if ($pkg.Manager -eq 'Winget') {
                    $wingetArgs = @('upgrade', '--id', $pkg.Id) + $WingetAcceptFlags
                    $exitCode = Invoke-WithRetry -PackageName $pkg.Name -ScriptBlock {
                        $out = & winget @wingetArgs 2>&1
                        $out | ForEach-Object { Write-UpdaterLog -Message "  $_" -Color 'DarkGray' }
                        return $LASTEXITCODE
                    }
                } elseif ($pkg.Manager -eq 'Chocolatey') {
                    $chocoArgs = @('upgrade', $pkg.Id, '-y')
                    $exitCode = Invoke-WithRetry -PackageName $pkg.Name -ScriptBlock {
                        $out = & choco @chocoArgs 2>&1
                        $out | ForEach-Object { Write-UpdaterLog -Message "  $_" -Color 'DarkGray' }
                        return $LASTEXITCODE
                    }
                }

                # Check success
                # Winget: 0 or -1978335189 (No applicable upgrade, already up to date)
                # Chocolatey: 0 usually
                if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
                    $status = 'Success'
                    Write-UpdaterLog -Message "Success." -Color "Green"
                } else {
                    $hex = '0x{0:X8}' -f ([int64]$exitCode -band 0xFFFFFFFF)
                    Write-UpdaterLog -Message "Failed. Exit code: $exitCode ($hex)." -Color "Red"
                }

            } catch {
                Write-UpdaterLog -Message "Exception during upgrade: $_" -Color "Red"
                $status = 'Error'
            }

            $results += [PSCustomObject]@{
                Manager = $pkg.Manager
                Name    = $pkg.Name
                Id      = $pkg.Id
                From    = $pkg.Version
                To      = $pkg.Available
                Result  = $status
            }
        }
    }

    # 6. Summary table
    Write-UpdaterLog -Message "--- Summary ---" -Color "Cyan"
    $tableLines = $results | Format-Table -Property Manager, Name, From, To, Result -AutoSize | Out-String
    foreach ($line in ($tableLines -split "`r?`n")) {
        if ($line.Trim()) {
            $color = if ($line -match 'Failed|Error') { 'Red' } elseif ($line -match 'Success') { 'Green' } else { 'Cyan' }
            Write-UpdaterLog -Message $line -Color $color
        }
    }
}

# ------------------------
# Main Execution Entry
# ------------------------
Assert-AdminPrivilege -ScriptParameters $PSBoundParameters
Invoke-PackageUpdate -ExclusionsFile $ExclusionsFile -Force:$Force -DryRun:$DryRun -SkipWinget:$SkipWinget -SkipChocolatey:$SkipChocolatey
