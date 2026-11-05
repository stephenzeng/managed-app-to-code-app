---
name: convert-managed-app-to-code-app
description: 'Convert an extracted Copilot Managed App built from Copilot Cowork App Skills into a Power Apps Code App: initialise, guide data-source setup, migrate data clients, fix runtime routes, build and publish.'
argument-hint: 'App display name and environment ID'
user-invocable: true
---

# Convert a Copilot Managed App to a Power Apps Code App

> **Disclaimer:** Created by Stephen Zeng. Not official Microsoft guidance. Managed apps, CLI commands and this process may change.

## Ground rules

- The app is already exported and extracted into the current working directory (the project root). Do not ask about export or extraction.
- Convert without redesigning the UI or losing existing behaviour.
- Never invent IDs, connection details, tenant details or data-source names. Never ask for passwords or tokens.
- Never replace real connector calls with sample data, and never hand-edit generated files.
- If a command fails, report it and stop. Do not continue as if it succeeded.

## Step 1 – Collect inputs

Before running anything, ask for both values in one form:

| Placeholder | Value |
| --- | --- |
| `<app-name>` | App display name |
| `<environment-id>` | Target Power Platform environment ID |

Wait until both are confirmed. Do not guess or use defaults. If the user needs help finding a value, explain how to find it. If they cancel, stop.

## Step 2 – Baseline build

```shell
bun install
bun run build
```

Note any existing failures so they are not mistaken for conversion problems.

## Step 3 – Initialise the Code App

If the project is not already initialised:

```shell
pa app init --display-name "<app-name>" --environment-id "<environment-id>"
```

If it is already initialised, check that the name and environment match. Do not overwrite them.

## Step 4 – Data sources (user runs these)

Ask: **Does this app connect to any data sources?**

- **No:** Skip to Step 5. If the code still calls connectors, point this out and ask the user to clarify.
- **Yes:** Ask the user to add each source from a terminal at the project root:

  ```shell
  pa app add data-source --connector <connector-id> [options]
  ```

  Options differ by connector and CLI version. Refer them to `pa app add data-source --help` and the [CLI reference](https://learn.microsoft.com/en-us/power-apps/developer/code-apps/reference/cli#pa-app-add-data-source). Do not run these commands yourself.

  Ask the user to type **continue** when finished. If they report a failure, help them fix it before moving on.

## Step 5 – Update runtime routing

Run:

```shell
bun install
bun run build
```

Then set `src/routes.tsx` to:

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

## Step 6 – Migrate data access

Only for apps with data sources:

1. Find every reference to the old root-level `generated/` client (imports, aliases, wrappers, config).
2. Read the new services and models in `src/generated/` and switch each call over. Map arguments, fields, IDs, response shapes, paging and errors. This is an API migration, not just an import rename.
3. Keep all existing reads, filters, sorting and writes. Refresh affected queries after writes.
4. Keep loading, empty, error and permission-denied states. Render connector text safely.
5. Remove dependencies only when they are proven unused.

Delete the root `generated/` folder only when nothing references it, the build passes and the user approves. Never delete `src/generated/`.

## Step 7 – Validate, build and publish

Run any configured tests, lint or type checks, then:

```shell
bun run build
pa app push
```

Publish automatically once the build passes, without asking for confirmation. Then tell the user to open the app from the Power Apps maker portal: https://make.powerapps.com

## Completion report

Summarise:

- What changed
- Checks run and their results
- Publishing status
- Any remaining manual steps or blockers
