# Azure Setup Read Me

This file contains the Microsoft Entra / Azure-side setup needed for the `MDCPP Custom` package.

Official Microsoft references:

- MDCPP samples repo:
  `https://github.com/microsoft/MDCPP-samples`
- Present Live sample README:
  `https://github.com/microsoft/MDCPP-samples/blob/main/present-live-integration-app-sample/README.md`

## App registration

1. Open Microsoft Entra admin center:
   `https://entra.microsoft.com/`
2. Go to:
   `Applications` -> `App registrations`
3. Create a new app registration or reuse an existing one
4. Supported account type:
   `Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)`

## Redirect URI

Add this SPA redirect URI:

```text
http://localhost:8080/desktop.html
```

If HTTPS is used later, also add:

```text
https://localhost:8080/desktop.html
```

## Required permissions for this Team Startpage prototype

Microsoft Graph delegated permissions:

- `User.Read`
- `Files.Read`
- `Files.ReadWrite`

Why:

- `User.Read` loads signed-in user details
- `Files.Read` loads files and previews
- `Files.ReadWrite` supports upload, create, delete, and editable Office attempts

## Admin consent

After adding permissions:

1. Open `API permissions`
2. Select `Grant admin consent`
3. Confirm it for the tenant

## Local `.env` value

Create a local `.env` file inside:

`present-live-integration-app-sample`

Add:

```text
ENTRA_APPID=[your Entra client ID]
```

## Important note

This upload package does not include `.env`.
That is intentional and safer for org upload/handoff.
