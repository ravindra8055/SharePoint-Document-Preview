# SharePoint Document Preview POC

A .NET 8 Blazor Server application designed to interface with a SharePoint Online document library. It displays all files contained within a specific folder in a tabular data grid. If SharePoint is unavailable, it provides a fallback option to upload a CSV file containing file metadata and preview URLs.

## Features

* **SharePoint Integration**: Connects to SharePoint Online using the official Microsoft Graph API SDK.
* **Entra ID Authentication**: Uses Microsoft Authentication Library (MSAL) for secure, credential-less authentication via Azure AD.
* **Data Grid**: Displays files with exactly three columns: File Name, File Size (Bytes), and Last Modified Date.
* **CSV Fallback**: Allows users to upload a CSV file containing file data if the SharePoint connection is unavailable.
* **Document Preview**: Includes an action button in the grid to open a document's preview URL in a new tab (populated via the CSV fallback).

## Prerequisites

* [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
* An Azure Active Directory (Entra ID) tenant.
* An App Registration in Entra ID with the following Application Permissions (Admin consent required):
  * `Sites.Read.All` or `Files.Read.All`

## Setup & Configuration

1. Navigate to the project directory:
   ```bash
   cd SharePointViewer
   ```

2. Open `appsettings.json` and replace the placeholders with your Entra ID App Registration details:
   ```json
   "AzureAd": {
     "Instance": "https://login.microsoftonline.com/",
     "Domain": "yourdomain.onmicrosoft.com",
     "TenantId": "YOUR_TENANT_ID",
     "ClientId": "YOUR_CLIENT_ID",
     "ClientSecret": "YOUR_CLIENT_SECRET"
   }
   ```

## Running the Application

Run the application using the .NET CLI:

```bash
dotnet run
```

Open your browser and navigate to the localhost URL provided in the terminal output.

## CSV Fallback Format

To use the CSV fallback feature, ensure your CSV file includes a header row and follows this exact column structure:

```csv
Name,Size,LastModified,PreviewUrl
Project_Proposal.pdf,1048576,2023-10-25T14:30:00Z,https://example.com/preview/1
Budget_2024.xlsx,512000,2023-11-01T09:15:00Z,https://example.com/preview/2
```
