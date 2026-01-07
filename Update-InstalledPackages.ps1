<#
.SYNOPSIS
    Updates installed packages using winget with improved logging and exclusions.

.DESCRIPTION
    This script checks for available upgrades using winget, filters them based on a JSON exclusion list,
    and performs upgrades. It produces a log file for each run.

.PARAMETER LogPath
    Optional path to a specific log file or directory. If a directory is provided, a timestamped log is created.

.PARAMETER ExclusionsFile
    Path to the JSON file containing package IDs to exclude.

.PARAMETER Force
    If specified, skips the confirmation prompt.

.EXAMPLE
    .\Update-InstalledPackages.ps1 -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath,

    [Parameter(Mandatory=$false)]
    [string]$ExclusionsFile,

    [switch]$Force
)

Add-Type -AssemblyName System.Windows.Forms

# ------------------------
# Configuration & Setup
# ------------------------

# Determine Script Directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) {
    $ScriptDir = "C:\Users\$env:USERNAME\Documents\chatgpt"
}

# Determine Log File Path
if (-not $LogPath) {
    $LogBaseDir = Join-Path $ScriptDir "logs"
    if (-not (Test-Path -Path $LogBaseDir)) {
        New-Item -Path $LogBaseDir -ItemType Directory -Force | Out-Null
    }
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:CurrentLogFile = Join-Path $LogBaseDir ("winget-upgrade-" + $TimeStamp + ".log")
} elseif (Test-Path -Path $LogPath -PathType Container) {
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:CurrentLogFile = Join-Path $LogPath ("winget-upgrade-" + $TimeStamp + ".log")
} else {
    $parent = Split-Path -Parent $LogPath
    if (-not (Test-Path -Path $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }
    $Script:CurrentLogFile = $LogPath
}

# Determine Exclusions File
if (-not $ExclusionsFile) {
    $ExclusionsFile = Join-Path $ScriptDir "winget-upgrade-exclusions.json"
}

$AcceptFlags = '--accept-package-agreements --accept-source-agreements'

# ------------------------
# Functions
# ------------------------

function Assert-AdminPrivilege {
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
        $argsList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $MyInvocation.MyCommand.Path)
        $BoundParameters.GetEnumerator() | ForEach-Object {
            if ($_.Value -is [switch] -and $_.Value) {
                $argsList += "-$($_.Key)"
            } elseif ($_.Value -isnot [switch]) {
                $argsList += "-$($_.Key)"
                $argsList += "`"$($_.Value)`""
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

function Write-WingetLog {
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
        "Winget Upgrade Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

function Get-WingetUpgrade {
    <#
    .SYNOPSIS
        Wraps the logic of trying JSON first, then failing back to Table parsing.
    #>
    Write-WingetLog -Message "Querying winget for available upgrades..." -Color "Cyan"

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
            Write-WingetLog -Message "Successfully parsed upgrades from winget JSON output." -Color "DarkGray"
            return $upgrades
        } else {
            throw "Failed to parse JSON from winget output."
        }
    } catch {
        if ($triedJson) {
            Write-WingetLog -Message ("Winget JSON attempt failed: {0}. Falling back to table parsing." -f $_.Exception.Message) -Color "Yellow"
        } else {
            Write-WingetLog -Message "Winget does not support JSON output (or failed). Falling back to table parsing." -Color "Yellow"
        }
    }

    # Fallback
    Write-WingetLog -Message "Parsing winget table output..." -Color "DarkGray"
    $raw = winget upgrade 2>&1
    Add-Content -Path $Script:CurrentLogFile -Value "`n===== RAW winget table output =====`n"
    return ConvertFrom-WingetTable -RawLines $raw
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
        Write-WingetLog -Message "Error internal parsing JSON: $_" -Color "Red"
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
        $pkgLines = $lines[($sepIndex + 1)..($lines.Length - 1)] | Where-Object { $_.Trim() -ne '' }

        $results = @()
        foreach ($line in $pkgLines) {
            # Basic parsing strategy: multiple spaces are column separators
            $l = $line -replace '\s+winget\s*$', '' # remove trailing 'winget' source if present
            $normalized = ($l -replace '\s{2,}', ' | ').Trim()
            $cols = $normalized -split '\s\|\s'

            if ($cols.Count -ge 2) {
                $name = $cols[0].Trim()
                # Handle Name [Id] format if present in first col
                if ($name -match '^(.*)\s\[(.*)\]$') {
                    $name = $matches[1]
                    $id = $matches[2]
                } else {
                    $id = $cols[1].Trim()
                }

                $results += [PSCustomObject]@{
                    Name      = $name
                    Id        = $id
                    Version   = if ($cols.Count -ge 3) { $cols[2].Trim() } else { $null }
                    Available = if ($cols.Count -ge 4) { $cols[3].Trim() } else { $null }
                    Source    = ''
                }
            }
        }
        return $results
    } catch {
        return $null
    }
}

function Invoke-WingetUpdate {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ExclusionsFile,

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    # 1. Load Exclusions
    $ExcludeIds = @()
    if ($ExclusionsFile -and (Test-Path $ExclusionsFile)) {
        try {
            $ExcludeIds = Get-Content -Path $ExclusionsFile -Raw | ConvertFrom-Json -ErrorAction Stop
            Write-WingetLog -Message "Loaded exclusions from $ExclusionsFile" -Color "DarkGray"
        } catch {
            Write-Warning "Could not read '$ExclusionsFile'. Error: $_"
        }
    }

    Write-WingetLog -Message "Starting upgrade session. Log: $Script:CurrentLogFile" -Color "Cyan"

    # 2. Get Upgrades
    $upgrades = Get-WingetUpgrade

    if (-not $upgrades -or $upgrades.Count -eq 0) {
        Write-WingetLog -Message "No upgradable packages detected." -Color "Green"
        return
    }

    Write-WingetLog -Message ("Found {0} potential upgrades." -f $upgrades.Count) -Color "Cyan"

    # 3. Apply Exclusions
    $toUpgrade = @()
    foreach ($u in $upgrades) {
        # Normalize ID for comparison (remove version tags sometimes in ID field e.g. "Vendor.App [Source]")
        $cleanId = ($u.Id -split '[\s\[]')[0]

        $isExcluded = $false
        foreach ($ex in $ExcludeIds) {
            if ($cleanId -ieq $ex -or $u.Id -ieq $ex) {
                $isExcluded = $true; break
            }
        }

        if (-not $isExcluded) {
            $toUpgrade += $u
        }
    }

    if ($toUpgrade.Count -eq 0) {
        Write-WingetLog -Message "All available upgrades are excluded." -Color "Yellow"
        return
    }

    # 4. Confirm
    Write-WingetLog -Message "Planned upgrades:" -Color "Cyan"
    $toUpgrade | ForEach-Object {
        Write-WingetLog -Message (" - {0} [{1}] ({2} -> {3})" -f $_.Name, $_.Id, $_.Version, $_.Available)
    }

    if (-not $Force) {
        if (-not (Show-GuiConfirmation)) {
            Write-WingetLog -Message "User canceled." -Color "Yellow"
            return
        }
    } else {
        Write-WingetLog -Message "Force enabled, proceeding..." -Color "Cyan"
    }

    # 5. Execute
    $results = @()
    foreach ($pkg in $toUpgrade) {
        Write-WingetLog -Message "Upgrading $($pkg.Name) [$($pkg.Id)]..." -Color "Magenta"

        if ($PSCmdlet.ShouldProcess($pkg.Name, "Winget Upgrade")) {
            $wingetArgs = @('upgrade', '--id', $pkg.Id) + ($AcceptFlags -split ' ')

            $output = ""
            $exitCode = 0
            try {
                $p = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -NoNewWindow -PassThru -Wait
                $exitCode = $p.ExitCode
            } catch {
                $output = $_.Exception.Message
                $exitCode = -1
            }

            # Check success (0 or 'No applicable upgrade' code)
            $status = 'Failed'
            if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
                $status = 'Success'
                Write-WingetLog -Message "Success." -Color "Green"
            } else {
                Write-WingetLog -Message "Failed. Exit code: $exitCode. Error: $output" -Color "Red"
            }

            $results += [PSCustomObject]@{
                Name = $pkg.Name
                Id = $pkg.Id
                Result = $status
            }
        }
    }

    # 6. Summary
    Write-WingetLog -Message "--- Summary ---" -Color "Cyan"
    $results | ForEach-Object {
        $color = if ($_.Result -eq 'Success') { 'Green' } else { 'Red' }
        Write-WingetLog -Message ("{0}: {1}" -f $_.Name, $_.Result) -Color $color
    }
}

# ------------------------
# Main Execution Entry
# ------------------------
Assert-AdminPrivilege
Invoke-WingetUpdate -ExclusionsFile $ExclusionsFile -Force:$Force
