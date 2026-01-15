# Roadmap

This document outlines the future development plans for the **Windows Software Package Updater**.

## Short Term Goals (v1.x)

- [ ] **Task Scheduler Integration**: Add a helper script or switch to easily register the updater as a scheduled task for fully automated background updates.
- [ ] **Email Notifications**: Add support for SMTP reporting in addition to Slack.
- [ ] **Enhanced Logging**: formatting logs as JSON for better machine parsing.
- [x] **Chocolatey Support**: Added support for `choco upgrade all` alongside `winget`.

## Medium Term Goals (v2.x)

- [ ] **Configuration UI**: A simple WPF or WinForms GUI to manage the `winget-upgrade-exclusions.json` file without editing text manually.
- [ ] **Rollback Capability**: If an update fails, attempt to restore the previous version (dependent on Winget capabilities).
- [ ] **Remote Management**: Ability to trigger updates on remote machines via PowerShell Remoting (WinRM).

## Long Term Goals (v3.x)

- [ ] **Centralized Dashboard**: A web-based or central dashboard to view update status across multiple machines.
- [ ] **Intune Integration**: Scripts packaged specifically for deployment via Microsoft Intune.
