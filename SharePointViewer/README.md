# SharePoint Document Preview POC

A .NET 8 Blazor Server application designed to interface with a SharePoint Online document library. It displays all files contained within a specific folder in a tabular data grid. If SharePoint is unavailable, it provides a fallback option to upload a CSV file containing file metadata and preview URLs.

## Features

* **SharePoint Integration**: Connects to SharePoint Online using the official Microsoft Graph API SDK.
* **Entra ID Authentication**: Uses Microsoft Authentication Library (MSAL) for secure, credential-less authentication via Azure AD.
* **Data Grid**: Displays files with exactly three columns: File Name, File Size (Bytes), and Last Modified Date.
* **CSV Fallback**: Allows users to upload a CSV file containing file data if the SharePoint connection is unavailable.
* **Document Preview**: Includes an action button in the grid to open a document's preview URL in a new tab or inline via an iframe.

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

## CSV Fallback Format

To use the CSV fallback feature, ensure your CSV file includes a header row and follows this exact column structure:

```csv
Name,Size,LastModified,PreviewUrl,EmbedUrl
Project_Proposal.pdf,1048576,2023-10-25T14:30:00Z,https://contoso.sharepoint.com/.../Doc.aspx?web=1,https://contoso.sharepoint.com/sites/MySite/_layouts/15/Doc.aspx?sourcedoc={doc-id}&action=embedview
Budget_2024.xlsx,512000,2023-11-01T09:15:00Z,https://contoso.sharepoint.com/.../Doc.aspx?web=1,https://contoso.sharepoint.com/sites/MySite/_layouts/15/Doc.aspx?sourcedoc={doc-id}&action=embedview
```

### ⚠️ Important: "SharePoint refused to connect" in Iframe

By default, SharePoint Online blocks external websites from embedding its pages inside an `iframe` to prevent clickjacking attacks (using `X-Frame-Options: SAMEORIGIN`). If you use a standard SharePoint URL in your CSV, the inline preview will fail to load.

**How to fix this:**
You must use SharePoint's specific **Embed URL** format for the `EmbedUrl` column. 

1. Go to your SharePoint document library in the browser.
2. Open the document (Word, Excel, PowerPoint, PDF).
3. Click **File** > **Share** > **Embed**.
4. Look at the provided Embed Code and copy the URL inside the `src="..."` attribute.
5. It will look something like this: `https://[tenant].sharepoint.com/sites/[site]/_layouts/15/Doc.aspx?sourcedoc={id}&action=embedview`

Using URLs with `&action=embedview` signals to SharePoint that the document is being embedded, and it will relax the framing restrictions (provided the user viewing the app is logged into their Microsoft 365 account in that browser).
