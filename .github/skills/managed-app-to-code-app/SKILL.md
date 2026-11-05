---
name: managed-app-to-code-app
description: 'Convert or migrate a Copilot Managed App exported from Copilot Cowork, App Skills, Copilot Code, or Copilot Studio App into a Power Apps Code App. Use for exporting a managed app, initializing a Code App, connecting a SharePoint list, migrating generated data clients, fixing Power Apps runtime routes, building, publishing, and validating the converted app.'
argument-hint: 'Project folder, app display name, environment ID, connection ID, SharePoint site URL, and list name'
user-invocable: true
---

# Convert a Copilot Managed App to a Power Apps Code App

## Scope

Convert an exported managed app without redesigning its UI or losing its existing
behavior. The established procedure is for an app created with App Skills in
Copilot Cowork and backed by a SharePoint list. Other origins and data sources,
including Dataverse, require checking their connector-specific CLI options and
generated APIs; do not assume the SharePoint command works unchanged.

**Disclaimer:** This is not official Microsoft guidance. Managed apps, CLI
commands, and the conversion process can change.

## Inputs and safety

1. Begin with a single intake form, using a structured question tool when
   available, before running any commands or beginning the conversion. Ask the
   user to provide every required command placeholder:

   | Command placeholder | Required user-provided value |
   | --- | --- |
   | `<app-name>` | App display name. |
   | `<environment-id>` | Target Power Platform environment ID. |
   | `<connection-id>` | Existing SharePoint connection ID in that environment. |
   | `<SharePoint list name>` | Exact SharePoint list name. |
   | `<SharePoint site URL>` | SharePoint site URL. |

   In the same intake, ask for the extracted project folder (or whether export
   is still needed) and whether the user wants local conversion only or also
   publishing. If values were already supplied in the invocation, present them
   for confirmation rather than asking the user to type them again.

   Wait until the user has supplied or explicitly confirmed every required
   value. Do not defer collecting a value until the step that uses it, infer
   values from the workspace, or substitute examples or defaults. If a value
   is missing, blank, unknown, or still a placeholder, ask for it and pause the
   conversion. If the user needs help locating a value, explain how to find it
   and wait for their answer. If the user declines or cancels intake, stop
   without running conversion commands.
2. Do not invent IDs, connection details, tenant details, or data-source names.
   Reuse existing project settings only after confirming the target with the user.
   Never request passwords or access tokens in chat.
3. Read the project's applicable instructions and inspect its Git status,
   dependency manifest, routing, data access, and existing Code App configuration.
   Preserve unrelated and uncommitted changes. Do not reinitialize an already
   configured app or add a duplicate data source without investigating first.
4. Follow repository restrictions. If migration requires a protected scaffold or
   configuration change, explain the conflict and request the appropriate
   authorization rather than silently bypassing those instructions.
5. Ask before publishing unless the user has explicitly authorized deployment to
   the confirmed environment. Ask separately before deleting the old generated
   client. Do not commit or push source-control changes unless requested.
6. Keep connection-specific settings in the app's supported configuration, not
   in this skill. Do not copy customer data, credentials, or tenant-specific
   configuration into reusable instructions.

## Procedure

Run commands in the extracted app's root. The following commands reproduce the
source guide's `pa app` CLI syntax. Every angle-bracket value must come from the
completed user intake above. Replace placeholders with those confirmed values
and quote arguments safely for the current shell. Before executing each command,
check that no unresolved placeholder remains. Never execute the examples with
literal placeholders.

### 1. Export and extract the app

If the source app has not been exported, ask the user to open it in Copilot Cowork
and request that it be packaged as a ZIP file, then download and extract it.
Do not claim to have performed export or browser actions without doing them.
For an already extracted app, continue from its existing folder.

### 2. Verify prerequisites and establish a baseline

Required:

- Node.js and Bun.
- Power Apps CLI with the guide's `pa app` commands.
- Visual Studio Code and a GitHub Copilot license.
- A Power Platform environment with Code Apps enabled.
- At least the Environment Maker role.
- An existing SharePoint connection in the target environment.

Check installed tools:

```shell
node --version
bun --version
pa --version
pa app --help
```

Check command-specific help for `init`, `add data-source`, and `push` before using
them. The guide calls this Power Apps CLI but uses the executable `pa`; do not
silently substitute `pac` or assume the command families are interchangeable.
If `pa app` is unavailable or flags differ, consult the installed tool's help and
current official Power Apps documentation to establish a supported equivalent.
If that cannot be established, report the blocker rather than inventing syntax.
Use the CLI's supported sign-in flow if authentication is needed.

Install dependencies and run the original build:

```shell
bun install
bun run build
```

Record baseline failures separately from conversion regressions. Do not proceed
past a failed command as though it succeeded.

### 3. Initialize the Power Apps Code App

For an app that is not already initialized:

```shell
pa app init --display-name "<app-name>" --environment-id "<environment-id>"
```

Inspect the resulting configuration to verify the app identity and environment.
For an already initialized project, verify these settings rather than overwriting
them. Keep existing application behavior and build settings wherever possible.

### 4. Add the SharePoint data source

Confirm that the connection exists in the same target environment, then run:

```shell
pa app add data-source --connector shared_sharepointonline --connection-id "<connection-id>" --table "<SharePoint list name>" --dataset "<SharePoint site URL>"
```

Answer **No** if prompted, "Are you using a connection reference instead of a
connection ID". If the CLI cannot accept interactive input, have the user run
this command or use a noninteractive option only if documented by its help.

This step generates models and services under `src/generated/` according to the
guide. Verify the actual generated location and exported types. Do not manually
fabricate or edit generated files.

### 5. Adapt runtime routing

Read `src/routes.tsx` and the existing router/base-path setup before editing.
The source guide replaces the single home route with:

```tsx
export const routes: AppRoute[] = [
  {
    path: "/appruntime/:appId/t/:tenantId/runtimeproxy/:runtimeId",
    element: <HomePage />,
    label: "Home",
    icon: HomeIcon,
  },
];
```

Use the project's existing imports, route types, exports, and metadata. Adapt
this pattern to the actual router rather than replacing a multi-page route table
with a single route. Preserve child routes, navigation, fallback behavior, and
local development access as appropriate; avoid duplicating a prefix already
handled by the router's basename. Do not edit a locked router provider.

Verify direct loading and in-app navigation under a concrete runtime URL matching
this pattern. A successful build alone does not establish correct routing.

### 6. Migrate data access to the Code App client

Find every reference to the old root-level `generated/` client, including imports,
path aliases, wrappers, and configuration references. Read the actual newly
generated service and model APIs in `src/generated/`; this is an API migration,
not just a string replacement of import paths.

- Replace managed-app client calls with the generated Code App services.
- Map arguments, field names, IDs, response shapes, paging, and error behavior
  using the generated declarations.
- Preserve existing reads, searches, filters, sorting, and all implemented writes.
- Keep server data in the app's existing query layer and invalidate or refresh
  affected queries after writes.
- Preserve loading, empty, error, and permission-denied states.
- Render connector text safely and validate URLs using existing helpers.
- Do not replace real connector calls with sample data to make a build pass.
- Check dependencies and configuration for obsolete managed-runtime requirements;
  remove only those proven unused, subject to repository permissions.

Do not remove `generated/` from the project root until every consumer is migrated,
no references remain, the new integration builds, and the user has approved its
deletion. Distinguish it carefully from `src/generated/`, which must remain.
If deletion is not approved, leave the unused directory and report that fact.

### 7. Validate, build, and publish

Run the existing focused tests, lint, or type-check commands relevant to the
changed code, if configured, then:

```shell
bun run build
```

Check the runtime route and the new SharePoint integration where authenticated
execution is available. Verify list loading and each existing write operation;
obtain permission before test writes to live data and clean up only test records
created for this validation.

Before publishing, confirm the destination environment and app identity. With
the user's authorization:

```shell
pa app push
```

Retain the command output for the launch step. Report command failures explicitly;
do not launch an app after a failed push. If publishing was not requested, stop
after local validation and give the remaining publish and launch steps.

### 8. Launch and verify in Power Apps

After `pa app push` succeeds, extract the complete app URL from its success
message, such as "App pushed successfully. You can play your app at ...".
Use the URL from this execution, not a hardcoded example, a previous deployment,
or a URL constructed from IDs. Preserve its full path and query string,
including `&` characters. Terminal visual wrapping does not split the URL.
This URL is a command result, not an additional user-provided intake value.

Validate that it is an absolute HTTPS Power Apps play URL (normally the exact
host `apps.powerapps.com` with a `/play/` path). If the CLI returns a different
host, confirm that it is the target environment's legitimate Power Apps host
before opening it. Treat the URL as data, never as executable shell text.

Automatically open the validated URL in the system's **default browser**.
On Windows, store the exact returned URL in the PowerShell string variable
`$appUrl`, then run:

```powershell
Start-Process -FilePath $appUrl
```

Use the operating system's equivalent default-browser opener on other platforms.
Do not substitute an embedded VS Code browser or navigate to the maker portal
when the push already returned a valid app URL.

If publishing succeeds but no usable URL is returned, report that explicitly;
do not invent a URL or republish just to obtain one. As a manual fallback, direct
the user to https://make.powerapps.com, select the confirmed target environment,
go to **Apps**, and launch the Code App. If the default-browser opener fails or
is unavailable, report the failure and provide the returned URL for manual launch.

Ask the user to complete sign-in and authenticated checks in the default browser
when those cannot be observed with available, authorized browser tooling.

Verify:

- The home screen loads at the runtime route, including after a refresh.
- Existing navigation and secondary routes still work.
- The expected SharePoint list loads through the new generated client.
- Existing create, update, and delete actions work when present and authorized.
- Errors and permissions are surfaced rather than hidden.

Opening the browser is not proof that the app loaded correctly. Do not claim
end-to-end success when only the build, CLI publish, or browser launch was verified.

## Completion report

Summarize changes, checks actually run and their outcomes, publishing status,
and any remaining manual steps or blockers. Distinguish local build success,
successful publishing, default-browser launch, and verified runtime behavior.
Do not include secrets or raw connector payloads.
