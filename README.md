# MDCPP Custom

This project is a Microsoft 365-connected Team Startpage prototype built from the Microsoft MDCPP sample base.

Base references:

- MDCPP samples repo:
  `https://github.com/microsoft/MDCPP-samples`
- Present Live sample README:
  `https://github.com/microsoft/MDCPP-samples/blob/main/present-live-integration-app-sample/README.md`

What we changed on top of the base sample:

- added a Teams-inspired start page UI
- added Microsoft Graph-based file listing
- changed Home to show recent files only
- moved folders into a dedicated `Folders` section in the left navigation
- added folder browsing inside the app
- added focused in-window file preview
- added `Edit in Microsoft 365`
- added an `Edit in app` attempt for supported Office files
- added favorites, sorting, searching, upload, create, and recycle-bin flows

> **IMPORTANT**:
> This handoff package is based on the MDCPP sample, but it is now a customized Team Startpage prototype and not just the original Present Live sample anymore.

## Additional Team Startpage notes

Before following the same base structure as the Microsoft sample below, keep these extra points in mind:

1. This upload package intentionally excludes local-only files such as `.env`, `node_modules`, `dist`, `lib`, and log files.
1. Home shows recent files only.
1. Folders are handled from the left-side `Folders` section.
1. In-window open is primarily a preview flow.
1. Full editing should be done through:

   ```text
   Edit in Microsoft 365
   ```

1. We also added an `Edit in app` attempt for some Office files, but Microsoft may still require embedded sign-in depending on browser, tenant, and session behavior.

## Present-live Sample Integration App

The present-live experience is currently available in the [MDCPP npm package](https://aka.ms/MDCPP-npm-package). This sample app demonstrates how to use the MDCPP npm package to integrate this experience.

The present-live experience allows meeting participants to interact and collaborate using PowerPoint Live directly in your collaboration app. For more about the present-live experience and how to implement it in your app, see [Integrate present-live experiences in the Microsoft 365 Document Collaboration Partner Program](https://learn.microsoft.com/microsoft-365/document-collaboration-partner-program/scenarios/present).

> **IMPORTANT**:
> You must be a member of the [Microsoft 365 Document Collaboration Partner Program](https://aka.ms/MDCPPwebsite) in order to access MDCPP services.

## Set up

Follow these instructions to configure a sample app in your test tenant.

1. Sign in to the Entra admin center with your test tenant admin account: https://entra.microsoft.com/
1. Expand the **Applications** tab and open the **App registrations** page.

    ![App registrations link on Microsoft Entra admin center's panel.](present-live-integration-app-sample/images/admin-center-registrations-tab.png)

1. Select **New registration** and give your sample app a name.
1. Select **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant)** under **Supported account types**.
1. Add an SPA redirect URI pointing to http://localhost/desktop.html as shown in the following screenshot.

   ![Register an application.](present-live-integration-app-sample/images/entra-new-app-registration.png)

1. Select **Register**.
1. Choose the **API permissions** tab on your sample app's registration page.
1. Select **Add a permission** > **SharePoint** > **Delegated permissions**.
1. Ensure that **AllSites.Read** and **MyFiles.Read** permissions are selected. Select **Add permissions**.
1. Select **Grant admin consent for \<tenant-name\>** (where \<tenant-name\> is the name of your test tenant) then choose **Yes** in the subsequent popup dialog. Your registration page should look similar to the following screenshot.

   ![Tenant application registered.](present-live-integration-app-sample/images/entra-tenant-app-registered.png)

1. Ensure that your test user account's OneDrive files contain at least one PowerPoint slide deck for testing.

   > **Note**: You can create a sample PowerPoint presentation file at https://m365.cloud.microsoft/create.

## Additional Azure / Entra changes for this custom version

Because this Team Startpage version uses Microsoft Graph for user and file operations beyond the original Present Live sample, we additionally granted these Microsoft Graph delegated permissions:

- `User.Read`
- `Files.Read`
- `Files.ReadWrite`

Why these were added:

- `User.Read` for signed-in user profile details
- `Files.Read` for listing and previewing files
- `Files.ReadWrite` for upload, create, delete, and editable Office attempts

This is the main Azure-side difference between the base sample and this customized Team Startpage version.

## Run the sample app locally

1. Fork this repo then clone your fork locally.
1. Navigate to and open the root directory of the sample **present-live-integration-app-sample**.
1. Create a new file named **.env** in the sample's root directory.
1. Open the .env file in your preferred code editor.
1. Add the following line to the .env file, replacing "[your Entra client ID]" with your Entra sample app's client ID. You can find the ID by navigating to the **Overview** tab in your app's registration page, listed under **Application (client) ID** in the **Essentials** dropdown menu near the top of the page.

   ```text
   ENTRA_APPID = [your Entra client ID]
   ```

1. In a terminal, run the following command from the sample's root directory to download dependencies.

   ```text
   yarn
   ```

   1. If you don't have yarn, you can download it by running the following command.

      ```text
      npm i -g yarn
      ```

   1. If you don't have Node.js, you can download it from https://nodejs.org/en/download.

1. Run the following command to launch the sample app. Ensure that port 8080 is free.

   ```text
   yarn start
   ```

1. In a private or incognito browser window, navigate to the following URL. To avoid issues with signing in, it's recommended to do this in a private or incognito browser window in case you're already signed in with a different Entra account in that browser.

   ```text
   http://localhost:8080/desktop.html
   ```

1. Sign in with your test Entra account. You may need to accept a **Permissions requested** dialog. After a few seconds, the file picker UI should load.
1. Select your test PowerPoint file from the file picker.
1. Choose **Present** at the bottom right.

A new browser tab should open that contains the attendee view. The initial browser tab will contain the presenter view. You may need to allow browser popups in order to see the new attendee tab.

However, for you to test out the present-live functionality fully, you'll need to support HTTPS. See the following section [Using HTTPS](#using-https) for details.

## Additional local behavior changes in this custom version

After sign-in, this custom Team Startpage version behaves differently from the original sample:

1. The main experience is a Team Startpage, not just the original picker-first Present Live path.
1. Home shows recent files only.
1. Folders are opened from the dedicated left-side folder browser.
1. Files can open in a focused in-window preview.
1. Full edit should be done with:

   ```text
   Edit in Microsoft 365
   ```

1. An `Edit in app` option is also available for supported Office files as a best-effort embedded editing attempt.

## Using HTTPS

The MDCPP endpoints require the host app to support HTTPS for security. Please ensure your app is served using HTTPS. There are many ways to achieve this depending on your setup. Here are two methods that aren't specific to this program that use open source or free software.

### Method 1: Ngrok

1. Install and setup [ngrok](https://ngrok.com/download).
1. Create an account on ngrok.com.
1. Run `ngrok config add-authtoken [your auth token]`.
1. Run `ngrok http 8080` and copy the Forwarding URL from terminal.
1. Go to the [Azure App registrations page](https://aka.ms/AppRegistrations/?referrer=https%3A%2F%2Fdev.onedrive.com) and add the Forwarding URL to the list of approved Redirect URLs.
1. To run the present-live test app on localhost, follow the instructions starting at step 8 in the [Run the sample app locally](#run-the-sample-app-locally) section, but instead navigating to `https://localhost:8080/desktop.html`.
1. Open the app using the Forwarding URL.

### Method 2: mkcert (No need to add Redirect URL to registered App)

1. Install [mkcert](https://github.com/FiloSottile/mkcert).
1. Run `mkcert localhost 127.0.0.1` to generate a certificate for the host you are using locally. The output should show the path to the generated cert and key.
1. Open the webpack.config.js file in this repo's root directory and add the [devServer](https://webpack.js.org/configuration/dev-server/#devserverserver) configs to use the generated files.

   ```text
   devServer: {
     server: {
       type: 'https',
       options: {
         ca: './path/to/server.pem',
         pfx: './path/to/server.pfx',
         key: './path/to/server.key',
         cert: './path/to/server.crt',
         passphrase: 'webpack-dev-server',
         requestCert: true,
       },
     },
   },
   ```

1. You can find the path to the CA by running `mkcert -CAROOT`.
1. Start the present-live test app using localhost, following the instructions starting at step 8 in the [Run the sample app locally](#run-the-sample-app-locally) section, but instead navigate to `https://localhost:8080/desktop.html`.

## Safe upload note

This handoff package intentionally excludes:

- `.env`
- `node_modules/`
- `dist/`
- `lib/`
- local log files

That keeps the upload package safer and cleaner for org handoff.
