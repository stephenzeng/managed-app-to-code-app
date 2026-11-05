# Convert a Copilot Managed App to a Power Apps Code App

A reusable GitHub Copilot skill for converting a Copilot Managed App to a Power Apps Code App.

## Please note

This is not official Microsoft guidance. It is a personal project created by Stephen Zeng.

Microsoft may change how Copilot Managed Apps are structured, and this skill may stop working as a result.

`kooka-issue-tracker` is included as an example app.

## How to install the skill

Copy the entire `convert-managed-app-to-code-app` folder under `skills`, including all files within it, to:

`%USERPROFILE%\.copilot\skills\`

For non-Windows operating systems, use the equivalent GitHub Copilot skills directory.

## How to use the skill

Before starting, make sure:

- Visual Studio Code is installed and GitHub Copilot is configured.
- The managed app source code is available on your local computer.

To convert an app:

1. Open the managed app folder in Visual Studio Code.
2. Open Copilot Chat and invoke the skill by typing `/convert-managed-app-to-code-app`.
3. Follow the prompts, providing the **app display name** and **target Power Platform environment ID** when requested.
4. Once the process completes successfully, the Power Apps Code App should be created in the specified Power Platform environment and be ready to launch.

## Troubleshooting

See the VS Code Agent Skills documentation for skill discovery and configuration.

If the skill is not listed in Copilot Chat:

- Reload the Visual Studio Code window.
- Confirm that skill discovery is enabled.
- Confirm that the `convert-managed-app-to-code-app` folder is in the correct skills directory.