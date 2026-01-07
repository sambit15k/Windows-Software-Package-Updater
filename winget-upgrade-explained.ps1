param(
    [switch]$Force
)

Add-Type -AssemblyName System.Windows.Forms

# ------------------------
# Configuration
# ------------------------
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) {
    $ScriptDir = "C:\Users\$env:USERNAME\Documents\chatgpt"
}
if (-not (Test-Path -Path $ScriptDir)) {
    New-Item -Path $ScriptDir -ItemType Directory -Force | Out-Null
}

$ExclusionsFile = Join-Path $ScriptDir "winget-upgrade-exclusions.json"
$ExcludeIds = @()

if (Test-Path $ExclusionsFile) {
    try {
        $ExcludeIds = Get-Content -Path $ExclusionsFile -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Warning "Could not read or parse '$ExclusionsFile'. No packages will be excluded. Error: $_"
    }
}

$TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = Join-Path $ScriptDir ("winget-upgrade-" + $TimeStamp + ".log")
$AcceptFlags = '--accept-package-agreements --accept-source-agreements'

# ------------------------
# Helpers
# ------------------------
function Ensure-RunAsAdmin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-Host "Not running as Administrator. Relaunching elevated..." -ForegroundColor Yellow
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        # Launch same shell (use pwsh if available), compatible with PS5.1
        if (Get-Command pwsh -ErrorAction SilentlyContinue) {
            $pwshExe = 'pwsh.exe'
        } else {
            $pwshExe = 'powershell.exe'
        }
        $psi.FileName = $pwshExe
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $MyInvocation.MyCommand.Path + '"'
        $psi.Verb = 'runas'
        try {
            [System.Diagnostics.Process]::Start($psi) | Out-Null
        } catch {
            Write-Host "Failed to relaunch elevated: $_" -ForegroundColor Red
        }
        Exit
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Color = 'Gray'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] $Message"
    Write-Host $line -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $line
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

function Get-WingetUpgradesFromJson {
    param(
        [Parameter(Mandatory=$true)][string]$RawText
    )

    try {
        # locate first JSON opening bracket: try '[' first then '{'
        $firstOpen = $RawText.IndexOf('[')
        if ($firstOpen -lt 0) { $firstOpen = $RawText.IndexOf('{') }
        if ($firstOpen -lt 0) { throw "No JSON start ('[' or '{') found in winget output." }

        $lastCloseArray = $RawText.LastIndexOf(']')
        $lastCloseObject = $RawText.LastIndexOf('}')
        if ($RawText[$firstOpen] -eq '[') {
            $lastClose = $lastCloseArray
        } else {
            $lastClose = $lastCloseObject
        }
        if ($lastClose -lt $firstOpen) { throw "Unable to locate matching JSON closing bracket." }

        $jsonCandidate = $RawText.Substring($firstOpen, ($lastClose - $firstOpen + 1))

        Add-Content -Path $LogFile -Value "`n===== EXTRACTED JSON CANDIDATE START =====`n"
        Add-Content -Path $LogFile -Value $jsonCandidate
        Add-Content -Path $LogFile -Value "`n===== EXTRACTED JSON CANDIDATE END =====`n"

        $parsed = $jsonCandidate | ConvertFrom-Json -ErrorAction Stop

        $upgrades = foreach ($item in $parsed) {
            $id = $null
            $name = $null
            $ver = $null
            $avail = $null

            if ($item.PSObject.Properties.Name -contains 'Id') { $id = $item.Id }
            elseif ($item.PackageId) { $id = $item.PackageId }

            if ($item.PSObject.Properties.Name -contains 'Name') { $name = $item.Name }
            elseif ($item.PackageName) { $name = $item.PackageName }

            if ($item.PSObject.Properties.Name -contains 'Version') { $ver = $item.Version }
            if ($item.PSObject.Properties.Name -contains 'AvailableVersion') { $avail = $item.AvailableVersion }

            if ($id -and $name) {
                [PSCustomObject]@{
                    Name = $name
                    Id = $id
                    Version = $ver
                    Available = $avail
                    Source = ($item.Source -as [string])
                }
            }
        }
        return $upgrades
    } catch {
        Write-Log -Message ("Error parsing winget JSON: {0}" -f $_) -Color "Yellow"
        return $null
    }
}

function Get-WingetUpgradesFromTable {
    param(
        [Parameter(Mandatory=$true)]$RawLines
    )

    try {
        $lines = $RawLines -split "`r?`n"
        $sepMatch = $lines | Select-String '^-{3,}' | Select-Object -First 1
        if (-not $sepMatch) {
            Write-Log -Message "Unable to parse winget table output (no separator found)." -Color "Red"
            return $null
        }

        $sepIndex = $sepMatch.LineNumber - 1
        $pkgLines = $lines[($sepIndex + 1)..($lines.Length - 1)] | Where-Object { $_.Trim() -ne '' }

        $upgrades = foreach ($line in $pkgLines) {
            $l = $line -replace '\s+winget\s*$', ''
            $normalized = ($l -replace '\s{2,}', ' | ').Trim()
            $cols = $normalized -split '\s\|\s'

            $name = $null; $id = $null; $ver = $null; $avail = $null

            if ($cols.Count -ge 1) {
                if ($cols[0] -match '\[([^\]]+)\]') {
                    $id = $matches[1].Trim()
                    $name = ($cols[0] -replace '\s*\[[^\]]+\]\s*', '').Trim()
                } else {
                    $name = $cols[0].Trim()
                }
            }

            if (-not $id -and $cols.Count -ge 2) { $id = $cols[1].Trim() }
            if ($cols.Count -ge 3) { $ver = $cols[2].Trim() }
            if ($cols.Count -ge 4) { $avail = $cols[3].Trim() }

            if ($avail -and ($avail -match '^\s*winget\s*$')) { $avail = $null }

            if ($id -and $name) {
                [PSCustomObject]@{
                    Name = $name
                    Id = $id
                    Version = $ver
                    Available = $avail
                    Source = ''
                }
            } else {
                Add-Content -Path $LogFile -Value ("UNPARSED LINE: {0}`n" -f $line)
            }
        }
        return $upgrades
    } catch {
        Write-Log -Message ("Error parsing winget table: {0}" -f $_) -Color "Yellow"
        return $null
    }
}

# ------------------------
# Start
# ------------------------
Ensure-RunAsAdmin

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Log -Message "Winget is not installed or not available in PATH." -Color "Red"
    Exit 1
}

Write-Log -Message "Starting winget upgrade session" -Color "Cyan"
$msg = "Log file: " + $LogFile
Write-Log -Message $msg -Color "DarkGray"
$msg = "Excluded IDs: " + ($ExcludeIds -join ', ')
Write-Log -Message $msg -Color "DarkGray"

# ------------------------
# Query winget for upgrades
# ------------------------
Write-Log -Message "Querying winget for available upgrades..." -Color "Cyan"
$upgrades = @()
$triedJson = $false

try {
    $rawLines = winget upgrade --output json 2>&1
    $rawText = $rawLines -join "`n"

    # Heuristic to check if winget supports --output json
    if ($rawText -match 'Unknown option|unrecognized|not a valid option|--output' -and -not ($rawText.TrimStart().StartsWith('[') -or $rawText.TrimStart().StartsWith('{'))) {
        throw "winget does not appear to support --output json."
    }

    $triedJson = $true
    Add-Content -Path $LogFile -Value "`n===== RAW winget JSON ATTEMPT START =====`n$rawText`n===== RAW winget JSON ATTEMPT END =====`n"
    $upgrades = Get-WingetUpgradesFromJson -RawText $rawText

    if ($upgrades) {
        Write-Log -Message "Successfully parsed upgrades from winget JSON output." -Color "DarkGray"
    } else {
        throw "Failed to parse JSON from winget output."
    }
} catch {
    if ($triedJson) {
        Write-Log -Message ("Winget JSON attempt failed: {0} Falling back to table parsing." -f $_.Exception.Message) -Color "Yellow"
    } else {
        Write-Log -Message "Winget does not support JSON output. Falling back to table parsing." -Color "Yellow"
    }
    $upgrades = @() # Ensure upgrades is empty for fallback
}

# Fallback to table parsing if JSON failed or was not supported
if ($upgrades.Count -eq 0) {
    Write-Log -Message "Parsing winget table output..." -Color "DarkGray"
    $raw = winget upgrade 2>&1
    Add-Content -Path $LogFile -Value "`n===== RAW winget table output START =====`n$raw`n===== RAW winget table output END =====`n"
    $upgrades = Get-WingetUpgradesFromTable -RawLines $raw

    if ($upgrades) {
        Write-Log -Message "Successfully parsed upgrades from winget table output." -Color "DarkGray"
    } else {
        Write-Log -Message "Failed to parse table output from winget." -Color "Red"
    }
}

if ($upgrades.Count -eq 0) {
    Write-Log -Message "No upgradable packages detected by winget." -Color "Yellow"
    Exit 0
}

$msg = ("Found {0} total upgradable package(s):" -f $upgrades.Count)
Write-Log -Message $msg -Color "Cyan"
foreach ($u in $upgrades) {
    $line = (" - {0} [{1}]  Current: {2}  Available: {3}" -f $u.Name, $u.Id, ($u.Version -as [string]), ($u.Available -as [string]))
    Write-Log -Message $line -Color "DarkGray"
}

# ------------------------
# Exclusions & planned list (Option A: attempt upgrade for any reported package with an Id)
# ------------------------
# Build the planned upgrade list: attempt upgrade for any package that has an Id and is not excluded
$toUpgrade = $upgrades | Where-Object {
    # must have Id
    if (-not $_.Id) { return $false }

    # normalize Id by taking the portion before any whitespace or '[' which often contains version tags
    $normId = ($_.Id -split '[\s\[]')[0]

    # check exclusion list (case-insensitive)
    foreach ($ex in $ExcludeIds) {
        if ($normId -ieq $ex -or $_.Id -ieq $ex) { return $false }
    }

    # keep the package (don't require Available)
    return $true
}

if (-not $toUpgrade -or $toUpgrade.Count -eq 0) {
    Write-Log -Message "After applying exclusions and filtering, no packages remain to upgrade. Exiting." -Color "Yellow"
    Exit 0
}

Write-Log -Message "The following packages WILL be upgraded (exclusions applied):" -Color "Cyan"
$index = 1
foreach ($p in $toUpgrade) {
    $line = ("{0}. {1} [{2}]  Current: {3}  Available: {4}" -f $index, $p.Name, $p.Id, ($p.Version -as [string]), ($p.Available -as [string]))
    Write-Log -Message $line -Color "Gray"
    $index++
}

# ------------------------
# Confirmation (GUI)
# ------------------------
if (-not $Force) {
    if (-not (Show-GuiConfirmation)) {
        Write-Log -Message "User canceled the upgrade via GUI. Exiting." -Color "Yellow"
        Exit 0
    }
} else {
    Write-Log -Message "Force mode enabled. Proceeding without confirmation..." -Color "Cyan"
}

# ------------------------
# Perform upgrades
# ------------------------
Write-Log -Message "Starting upgrades..." -Color "Green"
$Summary = foreach ($pkg in $toUpgrade) {
    $name = $pkg.Name
    $id = $pkg.Id
    $msg = ("Upgrading: {0} [{1}]" -f $name, $id)
    Write-Log -Message $msg -Color "Magenta"
    $wingetArgs = @('upgrade', '--id', $id) + ($AcceptFlags -split ' ')
    $global:LASTEXITCODE = $null
    try {
        $out = & winget @wingetArgs 2>&1
    } catch {
        $out = $_.Exception.Message
    }
    Add-Content -Path $LogFile -Value ("`n--- Output for {0} [{1}] START ---`n" -f $name, $id)
    Add-Content -Path $LogFile -Value $out
    Add-Content -Path $LogFile -Value ("`n--- Output for {0} [{1}] END ---`n" -f $name, $id)

    $exit = $LASTEXITCODE
    $success = $false

    # Primary check: Use winget's exit codes for success.
    # 0 is generic success. -1978335189 is WINGET_ERROR_NO_APPLICABLE_UPGRADE.
    if ($exit -eq 0 -or $exit -eq -1978335189) {
        $success = $true
    }
    # Secondary check: For older winget versions or ambiguous exit codes,
    # trust output text that explicitly indicates success or that no update is needed.
    elseif ($out -match 'Successfully (installed|upgraded)|already up to date|No (applicable update|updates found|available upgrade found)') {
        $success = $true
    }

    if ($success) {
        $msg = ("Upgrade succeeded: {0}" -f $name)
        Write-Log -Message $msg -Color "Green"
        [PSCustomObject]@{ Name = $name; Id = $id; Result = 'Success' }
    } else {
        $msg = ("Upgrade FAILED: {0} (ExitCode: {1})" -f $name, $exit)
        Write-Log -Message $msg -Color "Red"
        [PSCustomObject]@{ Name = $name; Id = $id; Result = 'Failed' }
    }
}

# ------------------------
# Final summary
# ------------------------
Write-Log -Message "Summary of results:" -Color "Green"
foreach ($s in $Summary) {
    if ($s.Result -eq 'Success') { $color = 'Green' } else { $color = 'Red' }
    $line = (" - {0} [{1}] => {2}" -f $s.Name, $s.Id, $s.Result)
    Write-Log -Message $line -Color $color
}

$succ = ($Summary | Where-Object { $_.Result -eq 'Success' }).Count
$fail = ($Summary | Where-Object { $_.Result -eq 'Failed' }).Count

$msg = ("All done. {0} succeeded, {1} failed." -f $succ, $fail)
Write-Log -Message $msg -Color "Cyan"
$msg = ("Full log saved at: {0}" -f $LogFile)
Write-Log -Message $msg -Color "Yellow"
