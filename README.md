# Windows Software Package Updater

[![Enterprise CI/CD](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-ci.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-ci.yml)
[![Publish Release](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-release.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-release.yml)
[![Publish Package](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-nuget-publish.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/pipeline-nuget-publish.yml)
[![CodeQL Security](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/security-codeql.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/security-codeql.yml)
[![Dependabot Updates](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/dependabot/dependabot-updates/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/dependabot/dependabot-updates)

A PowerShell utility to automate Windows package updates using **Winget** (Windows Package Manager) and **Chocolatey**, with support for exclusions, retry logic, log rotation, and confirmation dialogs.

## Features

- **Multi-Manager Support** — Updates packages from both `winget` and `chocolatey` (if installed)
- **Smart Parsing** — Attempts JSON output first, falls back to table parsing for older `winget` versions
- **JSONC Exclusions** — Exclude packages via a `.jsonc` config file with full comment support (`//` and `/* */`)
- **Retry Logic** — Automatically retries on transient network errors (timeouts, connection resets)
- **Log Rotation** — Auto-deletes log files older than a configurable number of days
- **Full Output Logging** — Installer output is captured line-by-line into the log file
- **Dry Run Mode** — Preview planned upgrades without modifying anything
- **GUI Confirmation** — Interactive dialog to review before upgrading (skippable with `-Force`)
- **Auto-Elevation** — Automatically relaunches with Administrator privileges if needed
- **Formatted Summary** — End-of-run table showing Manager / Name / From / To / Result

## Requirements

- Windows 10 or later
- PowerShell 5.1+ (PowerShell 7+ recommended)
- `winget` (Windows Package Manager) — optional but recommended
- `chocolatey` — optional
- Administrator privileges

## Files

| File | Description |
|---|---|
| `Update-InstalledPackages.ps1` | Main script |
| `winget-upgrade-exclusions.jsonc` | Package IDs to exclude from upgrades (JSONC with comment support) |

## Usage

### Basic

```powershell
# Run with GUI confirmation dialog
.\Update-InstalledPackages.ps1

# Skip confirmation prompt
.\Update-InstalledPackages.ps1 -Force

# Preview what would be upgraded — no changes made
.\Update-InstalledPackages.ps1 -DryRun
```

### Selective Manager

```powershell
# Winget only
.\Update-InstalledPackages.ps1 -SkipChocolatey -Force

# Chocolatey only
.\Update-InstalledPackages.ps1 -SkipWinget -Force
```

### Custom Paths & Log Retention

```powershell
# Custom exclusions file
.\Update-InstalledPackages.ps1 -ExclusionsFile C:\config\my-exclusions.jsonc

# Custom log directory
.\Update-InstalledPackages.ps1 -LogPath C:\logs\

# Keep only the last 14 days of logs
.\Update-InstalledPackages.ps1 -Force -KeepLogs 14
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Force` | Switch | — | Skip GUI confirmation prompt |
| `-DryRun` | Switch | — | List planned upgrades without performing them |
| `-SkipWinget` | Switch | — | Skip Winget upgrade checks |
| `-SkipChocolatey` | Switch | — | Skip Chocolatey upgrade checks |
| `-LogPath` | String | `.\logs\` | Path to a log file or directory |
| `-ExclusionsFile` | String | `.\winget-upgrade-exclusions.jsonc` | Path to exclusions file (`.jsonc` or `.json`) |
| `-KeepLogs` | Int | `30` | Days to retain log files (0 = keep forever) |

## Configuration

### Exclusions File

Edit **`winget-upgrade-exclusions.jsonc`** to exclude packages from upgrades.
The file supports `//` line comments and `/* */` block comments:

```jsonc
[
  // Wallpaper app — auto-updates reset custom settings
  "Microsoft.BingWallpaper",

  // Managed separately by IT — do not auto-update
  "Adobe.Acrobat.Pro",

  /* Epic launcher updates can break running games */
  "EpicGames.EpicGamesLauncher"
]
```

To temporarily disable all exclusions, comment out every entry:

```jsonc
[
  // "Microsoft.BingWallpaper",
  // "Adobe.Acrobat.Pro",
  // "EpicGames.EpicGamesLauncher"
]
```

> **Note:** A plain `.json` file is also supported as a fallback for backwards compatibility.

## Logging

Logs are saved to `.\logs\` by default, with one timestamped file per run:

```
logs\system-upgrade-20261007-020000.log
```

Each log contains:
- All packages detected as upgradable
- Excluded packages (individually listed)
- Full installer output for each package (line by line)
- Retry attempts with error codes
- End-of-run summary table

Log files older than `-KeepLogs` days (default: 30) are automatically deleted at the start of each run.

## How It Works

1. **Elevation Check** — Relaunches elevated if not already running as Administrator
2. **Log Rotation** — Removes log files older than `-KeepLogs` days
3. **Query Upgrades** — Fetches available upgrades from Winget and/or Chocolatey
4. **Parsing** — JSON output attempted first; falls back to table parsing
5. **Filtering** — Removes excluded packages from the upgrade list
6. **Confirmation** — Shows GUI dialog (skipped with `-Force` or `-DryRun`)
7. **Execution** — Upgrades each package with retry on transient network errors:
   - **Winget**: `winget upgrade --id <ID> --accept-package-agreements --accept-source-agreements`
   - **Chocolatey**: `choco upgrade <ID> -y`
8. **Summary** — Prints a formatted results table and saves to log

### Retry Logic

The script automatically retries on the following transient network exit codes:

| Exit Code | Hex | Meaning |
|---|---|---|
| -2147012894 | `0x80072EE2` | `WINHTTP_ERROR_TIMEOUT` |
| -2147012867 | `0x80072EFD` | `WINHTTP_ERROR_CONNECTION_ERROR` |
| -2147012866 | `0x80072EFE` | `WINHTTP_ERROR_CONNECTION_ABORTED` |
| -2147023293 | `0x80070643` | General transient installer failure |

Up to **3 retries** with a **15-second delay** between attempts.

### Sample Summary Output

```
Manager   Name              From          To            Result
-------   ----              ----          --            ------
Winget    Windows Terminal  1.24.12741.0  1.25.2733.0   Success
Winget    Git               2.46.0        2.47.0        Success
Choco     nodejs            20.0.0        22.0.0        Failed
```

## Troubleshooting

**Installer failed with exit code `0x80072EE2`**
- Network timeout downloading the package — the retry logic will handle this automatically (up to 3 attempts)
- If it persists, try downloading the installer manually and running `Add-AppxPackage` or `winget install --manifest`

**Script does not run / execution policy error**
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**"Not running as Administrator"**
- The script auto-elevates via UAC — approve the prompt when it appears

**Parsing errors in logs**
- Run `winget --version` to verify installation
- If winget is very old, update it from the Microsoft Store

## Integrations

### Scheduled Task (Automated Nightly Updates)

```powershell
$action  = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument '-NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Update-InstalledPackages.ps1" -Force -KeepLogs 30'
$trigger = New-ScheduledTaskTrigger -Daily -At '3:00AM'
Register-ScheduledTask -TaskName 'WindowsPackageUpdater' -Action $action -Trigger $trigger -RunLevel Highest
```

### Slack Notifications

To enable Slack notifications for pipeline status:
1. Create an **Incoming Webhook** in your Slack workspace and copy the URL
2. Go to **GitHub Repo Settings → Secrets and variables → Actions → New repository secret**
3. Name: `SLACK_WEBHOOK_URL` — Value: your webhook URL

### Slack App (Repo Events)

To receive notifications for Issues, PRs, and pushes:
1. Install the [GitHub for Slack](https://slack.github.com/) app
2. Run `/github subscribe owner/repo` in your Slack channel

## License

MIT License — Feel free to modify and distribute
