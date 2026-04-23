# Azure Setup Read Me

This file captures the Azure / Entra-specific setup for the `MDCPP Custom` package.

Base references:

- MDCPP samples repo:
  `https://github.com/microsoft/MDCPP-samples`
- Present Live sample README:
  `https://github.com/microsoft/MDCPP-samples/blob/main/present-live-integration-app-sample/README.md`

This custom package is based on that sample, but adds the Team Startpage-specific Microsoft Graph permissions and behavior notes listed below.

## Base sample Entra setup

Follow the same base Microsoft sample setup flow:

1. Open Microsoft Entra admin center:
   `https://entra.microsoft.com/`
2. Go to:
   `Applications` -> `App registrations`
3. Create a new app registration or reuse an existing one
4. Supported account type:
   `Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)`

## Redirect URI

Use the same SPA redirect URI for the local app:

```text
http://localhost:8080/desktop.html
```

If HTTPS is used later, also add:

```text
https://localhost:8080/desktop.html
```

## Base sample permissions

From the original MDCPP sample setup, the documented SharePoint delegated permissions are:

- `AllSites.Read`
- `MyFiles.Read`

## Additional permissions granted for this custom Team Startpage version

Because this version also uses Microsoft Graph for user and file operations, the following additional Microsoft Graph delegated permissions were granted:

- `User.Read`
- `Files.Read`
- `Files.ReadWrite`

Why these additional permissions are needed:

- `User.Read` loads signed-in user information
- `Files.Read` loads file listings and in-window previews
- `Files.ReadWrite` supports upload, create, delete, and editable Office attempts

## Admin consent

After adding the permissions:

1. Open `API permissions`
2. Select `Grant admin consent`
3. Confirm the action for the tenant

## Local `.env` value

Create a local `.env` file inside:

`present-live-integration-app-sample`

Add:

```text
ENTRA_APPID=[your Entra client ID]
```

## Important note about `.env`

The upload package does **not** include `.env`.

That is intentional because:

- it is environment-specific
- it should be created by the receiving team locally
- it is safer not to upload local environment values in the shared package

## Current Team Startpage auth behavior

This custom version currently behaves like this:

- the app shows an explicit sign-in screen first
- sign-in uses an interactive Microsoft login flow
- after sign-in, Microsoft Graph tokens are requested silently where possible

## Editing note

This package supports:

- stable in-window preview
- `Edit in Microsoft 365`
- a best-effort `Edit in app` attempt for supported Office files

Important limitation:

- Microsoft may still request embedded sign-in in the iframe depending on browser session, tenant configuration, and Microsoft 365 host behavior

## Safe upload reminder

This handoff package intentionally excludes:

- `.env`
- `node_modules/`
- `dist/`
- `lib/`
- local log files

That keeps the upload set cleaner and safer for organizational handoff.
