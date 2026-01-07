# Windows Software Package Updater

[![Enterprise CI/CD](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/ci.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/ci.yml)
[![Publish Release](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/release.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/release.yml)
[![Publish Package](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/publish-package.yml/badge.svg)](https://github.com/sambit15k/Windows-Software-Package-Updater/actions/workflows/publish-package.yml)


A PowerShell utility to automate Windows package updates using `winget` (Windows Package Manager) with support for exclusions, logging, and confirmation dialogs.

## Features

- **Automated Updates**: Queries `winget` for available package upgrades and performs them automatically
- **Smart Parsing**: Detects `winget` version and uses JSON output (newer versions) with fallback to table parsing (older versions)
- **Exclusion Management**: Specify packages to exclude from upgrades via JSON configuration
- **GUI Confirmation**: Interactive dialog to review and confirm before upgrading
- **Force Mode**: Skip confirmation with `-Force` flag for automated deployments
- **Auto-Elevation**: Automatically runs with administrator privileges
- **Comprehensive Logging**: Detailed logs with timestamps for each upgrade attempt
- **Error Handling**: Robust error detection with multiple success criteria

## Requirements

- Windows 10 or later
- PowerShell 5.1+
- `winget` installed (Windows Package Manager)
- Administrator privileges

## Files

- **Update-InstalledPackages.ps1** - Main script with modular code, external configuration, and advanced features (recommended)
- **winget-upgrade-exclusions.json** - List of package IDs to exclude from upgrades
- **winget-upgrade-raw.json** - Sample reference of winget help output

## Usage

### Basic Usage
```powershell
# Run with confirmation dialog
.\Update-InstalledPackages.ps1

# Run with Force mode (skip confirmation)
.\Update-InstalledPackages.ps1 -Force

# Run silently without user interaction
.\Update-InstalledPackages.ps1 -Silent
```

### Advanced Usage
```powershell
# Use custom exclusions file
.\Update-InstalledPackages.ps1 -ExclusionsFile C:\temp\my-exclusions.json

# Use custom log directory
.\Update-InstalledPackages.ps1 -LogDir C:\temp\logs
```

## Configuration

### Exclusions File
Edit **winget-upgrade-exclusions.json** to exclude packages:
```json
[
    "Adobe.Acrobat.Reader.64-bit",
    "AOMEI.PartitionAssistant"
]
```

## Logging

Logs are saved automatically to:
- Default: `%USERPROFILE%\Documents\chatgpt\logs\winget-upgrade-YYYYMMDD-HHMMSS.log`
- Log file path is displayed when the script runs
- Organized in a dedicated `logs` directory for easy management

Logs contain:
- All packages detected as upgradable
- Excluded packages
- Upgrade attempts and results
- Raw `winget` output for debugging

## How It Works

1. **Elevation Check**: Ensures script runs as Administrator
2. **Winget Query**: Requests list of available upgrades
3. **Parsing**: Attempts JSON parsing first, falls back to table parsing if needed
4. **Filtering**: Removes excluded packages from upgrade list
5. **Confirmation**: Shows GUI dialog with packages to upgrade (can be skipped with `-Force`)
6. **Execution**: Upgrades each package individually with proper flags
7. **Reporting**: Summarizes results and saves to log file

## Exit Codes

- `0` - Successful completion
- `1` - Winget not installed or error occurred

## Troubleshooting

**"Winget is not installed"**
- Install Windows Package Manager from Microsoft Store or GitHub

**"Not running as Administrator"**
- Script will auto-elevate, but ensure you approve the UAC prompt

**Parsing errors in logs**
- Check if your `winget` version is compatible
- Run `winget --version` to verify installation

## Available Parameters

The script supports the following command-line parameters:

- `-Force` - Skip GUI confirmation and proceed automatically
- `-Silent` - Alias for `-Force`
- `-ExclusionsFile <path>` - Custom path to exclusions JSON file
- `-LogDir <path>` - Custom directory for log files

## Integrations

### Slack Notifications
To enable Slack notifications for pipeline status:
1.  **Incoming Webhook**: Create an Incoming Webhook in your Slack Workspace and copy the URL.
2.  **Repo Secret**: Go to GitHub Repo Settings > Secrets and variables > Actions > New repository secret.
3.  **Name**: `SLACK_WEBHOOK_URL`
4.  **Value**: Paste your webhook URL.

### Slack App Integration
To receive notifications for all repo events (Issues, PRs, etc.):
1.  Install the [GitHub for Slack](https://slack.github.com/) app.
2.  Run `/github subscribe owner/repo` in your Slack channel.

## License

MIT License - Feel free to modify and distribute