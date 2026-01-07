# Windows Software Package Updater

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

- **upgrade-script.ps1** - Main script with hardcoded exclusions (simpler, single-file solution)
- **winget-upgrade-explained.ps1** - Enhanced version with modular code and external configuration (recommended)
- **winget-upgrade-exclusions.json** - List of package IDs to exclude from upgrades
- **winget-upgrade-raw.json** - Sample reference of winget help output

## Usage

### Basic Usage
```powershell
# Run with confirmation dialog
.\upgrade-script.ps1

# Or use the enhanced version
.\winget-upgrade-explained.ps1
```

### Force Mode (Skip Confirmation)
```powershell
.\upgrade-script.ps1 -Force
.\winget-upgrade-explained.ps1 -Force
```

## Configuration

### Upgrade Script
Edit the `$ExcludeIds` array in **upgrade-script.ps1** to exclude packages:
```powershell
$ExcludeIds = @(
    'Adobe.Acrobat.Reader.64-bit',
    'AOMEI.PartitionAssistant'
)
```

### Enhanced Version
Edit **winget-upgrade-exclusions.json** to exclude packages:
```json
[
    "Adobe.Acrobat.Reader.64-bit",
    "AOMEI.PartitionAssistant"
]
```

## Logging

Logs are saved automatically to:
- Default: `%USERPROFILE%\Documents\chatgpt\winget-upgrade-YYYYMMDD-HHMMSS.log`
- Log file path is displayed when the script runs

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

## Which Version to Use?

- **upgrade-script.ps1**: Simple, standalone, all-in-one solution
- **winget-upgrade-explained.ps1**: Better organized, external configuration, easier to maintain

Both work identically from a user perspective; choose based on your preference.

## License

MIT License - Feel free to modify and distribute