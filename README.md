# MDCPP Custom

This folder is the clean upload package for the Team Startpage prototype built from the Microsoft MDCPP sample.

Official Microsoft references:

- MDCPP samples repo:
  `https://github.com/microsoft/MDCPP-samples`
- Present Live sample README:
  `https://github.com/microsoft/MDCPP-samples/blob/main/present-live-integration-app-sample/README.md`

## What this package is

This package contains only the files needed to understand, review, and run the customized prototype safely inside another repository or organization workspace.

Main idea:

- Microsoft 365 sign-in
- Teams-inspired start page UI
- recent-files-first home view
- folder browsing from the left panel
- in-window file preview
- edit handoff to Microsoft 365

## Package structure

This upload package is intentionally simplified.

Main app folder:

`present-live-integration-app-sample`

Important source files:

- `present-live-integration-app-sample/src/desktop.ts`
- `present-live-integration-app-sample/src/aadauth.ts`
- `present-live-integration-app-sample/src/spoPicker.ts`
- `present-live-integration-app-sample/src/static/desktop.html`

Additional setup guide:

- `AZURE_SETUP_README.md`

## What is included

Included because they are needed:

- source code in `src/`
- static assets in `images/`
- project config files like `package.json`, `tsconfig.json`, `webpack.config.js`, `webpack.development.js`
- lock/config files like `.gitignore`, `.npmrc`, `.prettierrc`, `yarn.lock`

## What is intentionally NOT included

Removed from this handoff package:

- `.env`
- `node_modules/`
- `dist/`
- `lib/`
- local log files
- extra working notes
- duplicate or nonessential outer documentation
- support / conduct / security / license metadata files

## Why `.env` is not included

`.env` usually contains local environment-specific values such as the Entra app ID.
That file should be created by the receiving team in their own environment.

Required local value:

```text
ENTRA_APPID=[your Entra client ID]
```

## How to run locally

1. Open the folder `present-live-integration-app-sample`
2. Create a `.env` file
3. Add:

```text
ENTRA_APPID=[your Entra client ID]
```

4. Install dependencies:

```text
npm install --legacy-peer-deps
```

5. Start the app:

```text
npm run start
```

6. Open:

```text
http://localhost:8080/desktop.html
```

## Current product behavior

- Home shows recent files only
- Folders are browsed from the `Folders` section in the left panel
- Files open in a focused in-window preview
- Full editing should be done through `Edit in Microsoft 365`
- There is also an `Edit in app` attempt for supported Office files, but Microsoft may still require an embedded sign-in depending on tenant/session/browser behavior

## Notes for reviewers

This package is meant to be the safe upload version.
It is intentionally smaller and cleaner than the full working folder.
