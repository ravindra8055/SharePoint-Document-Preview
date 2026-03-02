# SharePoint Document Preview POC

A .NET 8 Blazor Server application designed to interface with a SharePoint Online document library. It displays all files contained within a specific folder in a tabular data grid. If SharePoint is unavailable, it provides a fallback option to upload a CSV file containing file metadata and preview URLs.

## Features

* **SharePoint Integration**: Connects to SharePoint Online using the official Microsoft Graph API SDK.
* **Entra ID Authentication**: Uses Microsoft Authentication Library (MSAL) for secure, credential-less authentication via Azure AD.
* **Data Grid**: Displays files with exactly three columns: File Name, File Size (Bytes), and Last Modified Date.
* **CSV Fallback**: Allows users to upload a CSV file containing file data if the SharePoint connection is unavailable.
* **Document Preview**: Includes an action button in the grid to open a document's preview URL in a new tab or inline via an iframe.
* **Secure Inline Previews**: Uses the Microsoft Graph API `preview` endpoint to generate short-lived, embeddable URLs for SharePoint files, bypassing standard SharePoint iframe restrictions.

## Prerequisites

* [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
* An Azure Active Directory (Entra ID) tenant.
* An App Registration in Entra ID with the following Application Permissions (Admin consent required):
  * `Sites.Read.All` or `Files.Read.All`

## Setup & Configuration

1. Open `appsettings.json` and replace the placeholders with your Entra ID App Registration details. You can use either `ClientSecret` or `UsernamePassword` for the `AuthType`.

## Running the Application

Run the application using the .NET CLI:

```bash
dotnet run
```

Open your browser and navigate to the localhost URL provided in the terminal output.

## IIS Deployment Troubleshooting

If you publish this application to IIS and encounter the error: **"The Requested page cannot be accessed because the related configuration data for the page is invalid"**, it means IIS does not recognize the `AspNetCoreModuleV2` module in your `web.config` file.

**To fix this:**
1. Download and install the **.NET 8 Hosting Bundle** on the Windows Server running IIS. You can download it from the official Microsoft .NET 8 download page (look for "ASP.NET Core Runtime" -> "Hosting Bundle").
2. After installation, open an Administrator Command Prompt and run `iisreset` to restart IIS and load the new module.
3. Ensure your IIS Application Pool for this site is set to **"No Managed Code"** (since .NET Core/.NET 8 manages its own process).

## CSV Fallback Format

To use the CSV fallback feature, ensure your CSV file includes a header row and follows this exact column structure:

```csv
Name,Size,LastModified,PreviewUrl,EmbedUrl
Project_Proposal.pdf,1048576,2023-10-25T14:30:00Z,https://contoso.sharepoint.com/.../Doc.aspx?web=1,https://contoso.sharepoint.com/sites/MySite/_layouts/15/Doc.aspx?sourcedoc={doc-id}&action=embedview
Budget_2024.xlsx,512000,2023-11-01T09:15:00Z,https://contoso.sharepoint.com/.../Doc.aspx?web=1,https://contoso.sharepoint.com/sites/MySite/_layouts/15/Doc.aspx?sourcedoc={doc-id}&action=embedview
```

### ⚠️ Important: "SharePoint refused to connect" in Iframe

If you are getting a "refused to connect" error in the iframe, it is almost certainly caused by **Authentication Redirects and Third-Party Cookie Blocking**.

*(Note: The SharePoint "HTML Field Security" setting is actually for allowing external sites to be embedded INSIDE SharePoint, not the other way around. It will not fix this issue).*

**The Root Cause:**
When the iframe tries to load the SharePoint `EmbedUrl`, it needs to know who you are. If your browser blocks **Third-Party Cookies** (which is the default in Incognito mode, Safari, and increasingly in Chrome/Edge), the iframe cannot read your SharePoint login cookie. 
Because it thinks you aren't logged in, SharePoint redirects the iframe to the Microsoft Login page (`login.microsoftonline.com`). **Microsoft strictly blocks its login page from being loaded inside an iframe** to prevent credential theft. The browser blocks the login page, resulting in the "refused to connect" error.

**How to fix this:**

**Solution 1: Enable Third-Party Cookies in your Browser**
For the iframe to authenticate you silently, you must allow third-party cookies for the site.
* **Chrome**: Go to Settings > Privacy and security > Third-party cookies > Select "Allow third-party cookies" (or add your SharePoint tenant and your app's domain to the allowed list).
* **Edge**: Go to Settings > Cookies and site permissions > Manage and delete cookies and site data > Turn off "Block third-party cookies".

**Solution 2: Use an Anonymous "Anyone with the link" URL**
If your organization allows anonymous sharing, generate an "Anyone with the link can view" link for the document. Open that link in an Incognito window, click "Embed", and use *that* URL in your CSV. Because it doesn't require login, it won't redirect to the blocked Microsoft login page.
