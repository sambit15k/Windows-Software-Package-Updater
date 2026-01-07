# CI/CD Pipelines

This project uses **GitHub Actions** to ensure code quality, security, and automated releases.

## Workflows

### 1. Enterprise CI/CD (`pipeline-ci.yml`)
**Triggers:** Push to `main`, `develop` or Pull Requests.

*   **PSScriptAnalyzer & Security**:
    *   Runs `PSScriptAnalyzer` to lint PowerShell code.
    *   Uploads SARIF results to GitHub Security tab.
*   **Syntax & Integrity Check**:
    *   Verifies the script syntax using `System.Management.Automation.Language.Parser`.
*   **Build & Publish Artifact**:
    *   Packages the script and README.md.
    *   Publishes a NuGet package to GitHub Packages (GPR) on `main`.
*   **Slack Notification**: Sends a success/failure alert to Slack using a pinned secure action `rtCamp/action-slack-notify`.

### 2. Publish Release (`pipeline-release.yml`)
**Triggers:** Push of a tag starting with `v*` (e.g., `v1.0.0`).

*   **Package Release Assets**: Zips the script and config files.
*   **Create GitHub Release**: Creates a release entry on GitHub and attaches the ZIP file and script.
*   **Slack Notification**: Alerts that a new release has been deployed.

### 3. Publish NuGet Package (`pipeline-nuget-publish.yml`)
**Triggers:** Push of a tag starting with `v*`.

*   **Setup NuGet**: Configures the NuGet CLI.
*   **Update Version**: Syncs the `.nuspec` version with the git tag.
*   **Pack & Push**: Pushes the `.nupkg` to GitHub Packages.

## Security

*   **Action Pinning**: Third-party actions like `slack-notify` are pinned to specific commit hashes (e.g., `e31e87e...`) to prevent supply chain attacks.
*   **Permissions**: Least-privilege permissions are defined at the job level.
