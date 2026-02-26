<#
.SYNOPSIS
    Updates installed packages using Winget and Chocolatey with improved logging and exclusions.

.DESCRIPTION
    This script checks for available upgrades using both Winget and Chocolatey (if installed).
    It filters them based on a JSON exclusion list and performs upgrades. 
    It produces a unified log file for each run.

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
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:CurrentLogFile = Join-Path $LogPath ("system-upgrade-" + $TimeStamp + ".log")
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

$WingetAcceptFlags = '--accept-package-agreements --accept-source-agreements'

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
# Winget Functions
# ------------------------

function Get-WingetUpgrade {
    <#
    .SYNOPSIS
        Wraps the logic of trying JSON first, then failing back to Table parsing.
    #>
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
        [switch]$Force
    )

    # 1. Load Exclusions
    $ExcludeIds = @()
    if ($ExclusionsFile -and (Test-Path $ExclusionsFile)) {
        try {
            $ExcludeIds = Get-Content -Path $ExclusionsFile -Raw | ConvertFrom-Json -ErrorAction Stop
            Write-UpdaterLog -Message "Loaded exclusions from $ExclusionsFile" -Color "DarkGray"
        } catch {
            Write-Warning "Could not read '$ExclusionsFile'. Error: $_"
        }
    }

    Write-UpdaterLog -Message "Starting system upgrade session. Log: $Script:CurrentLogFile" -Color "Cyan"

    # 2. Get Upgrades from all sources
    $allUpgrades = @()
    
    # Winget
    $wingetUpgrades = Get-WingetUpgrade
    if ($wingetUpgrades) { $allUpgrades += $wingetUpgrades }
    
    # Chocolatey
    $chocoUpgrades = Get-ChocolateyUpgrade
    if ($chocoUpgrades) { $allUpgrades += $chocoUpgrades }

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

        if (-not $isExcluded) {
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
            
            $output = ""
            $exitCode = 0
            $status = 'Failed'

            try {
                if ($pkg.Manager -eq 'Winget') {
                    $wingetArgs = @('upgrade', '--id', $pkg.Id) + ($WingetAcceptFlags -split ' ')
                    $p = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -NoNewWindow -PassThru -Wait
                    $exitCode = $p.ExitCode
                } elseif ($pkg.Manager -eq 'Chocolatey') {
                    # choco upgrade <id> -y
                    $chocoArgs = @('upgrade', $pkg.Id, '-y')
                    $p = Start-Process -FilePath "choco" -ArgumentList $chocoArgs -NoNewWindow -PassThru -Wait
                    $exitCode = $p.ExitCode
                }

                # Check success 
                # Winget: 0 or -1978335189 (No applicable upgrade, sometimes happens if already done)
                # Chocolatey: 0 usually
                if ($exitCode -eq 0 -or $exitCode -eq -1978335189) {
                    $status = 'Success'
                    Write-UpdaterLog -Message "Success." -Color "Green"
                } else {
                    Write-UpdaterLog -Message "Failed. Exit code: $exitCode." -Color "Red"
                }

            } catch {
                $output = $_.Exception.Message
                Write-UpdaterLog -Message "Exception during upgrade: $output" -Color "Red"
                $status = 'Error'
            }

            $results += [PSCustomObject]@{
                Name = $pkg.Name
                Id = $pkg.Id
                Manager = $pkg.Manager
                Result = $status
            }
        }
    }

    # 6. Summary
    Write-UpdaterLog -Message "--- Summary ---" -Color "Cyan"
    $results | ForEach-Object {
        $color = if ($_.Result -eq 'Success') { 'Green' } else { 'Red' }
        Write-UpdaterLog -Message ("[{0}] {1}: {2}" -f $_.Manager, $_.Name, $_.Result) -Color $color
    }
}

# ------------------------
# Main Execution Entry
# ------------------------
Assert-AdminPrivilege
Invoke-PackageUpdate -ExclusionsFile $ExclusionsFile -Force:$Force
