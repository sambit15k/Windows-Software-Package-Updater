# Welcome to the Windows Software Package Updater Wiki

The **Windows Software Package Updater** is a robust PowerShell utility designed to automate the patching of Windows software using `winget`. It is built for system administrators and power users who need control, logging, and notifications for their patch management process.

## Navigation

*   **[Roadmap](Roadmap.md)**: See what's planned for future releases.
*   **[CI/CD Pipelines](CI-CD-Pipelines.md)**: Understand how our automated build and release process works.
*   **[Troubleshooting](Troubleshooting.md)**: Common errors and how to resolve them.

## Quick Start

1.  Clone the repository.
2.  Run `.\Update-InstalledPackages.ps1`.
3.  Follow the interactive prompts or use `-Force` for silent operation.

## Key Features

*   **Winget Integration**: Leverages Microsoft's official Windows Package Manager.
*   **Smart Exclusions**: Prevent specific sensitive packages from auto-updating.
*   **Enterprise Reporting**: Get Slack notifications for update status.
*   **Security**: Signed commits and pinned GitHub Actions for supply chain security.
