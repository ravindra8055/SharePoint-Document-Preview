using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SharePointViewer.Models;

namespace SharePointViewer.Services;

public class SharePointService
{
    private readonly GraphServiceClient _graphClient;

    public SharePointService(IConfiguration configuration)
    {
        var authType = configuration["SharePointAuth:AuthType"];
        var tenantId = configuration["SharePointAuth:TenantId"];
        var clientId = configuration["SharePointAuth:ClientId"];
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        if (string.Equals(authType, "UsernamePassword", StringComparison.OrdinalIgnoreCase))
        {
            var username = configuration["SharePointAuth:Username"];
            var password = configuration["SharePointAuth:Password"];

            var options = new UsernamePasswordCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            var usernamePasswordCredential = new UsernamePasswordCredential(
                username, password, tenantId, clientId, options);

            _graphClient = new GraphServiceClient(usernamePasswordCredential, scopes);
        }
        else
        {
            var clientSecret = configuration["SharePointAuth:ClientSecret"];

            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            _graphClient = new GraphServiceClient(clientSecretCredential, scopes);
        }
    }

    public async Task<List<SharePointFile>> GetFilesInFolderAsync(string folderUrl)
    {
        var files = new List<SharePointFile>();

        try
        {
            // Parse the URL to get the hostname, site path, and folder path
            var uri = new Uri(folderUrl);
            var hostname = uri.Host;
            
            // Example URL: https://contoso.sharepoint.com/sites/MySite/Shared%20Documents/MyFolder
            var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            
            if (segments.Length < 3)
                throw new ArgumentException("Invalid SharePoint folder URL format. Expected format: https://[tenant].sharepoint.com/sites/[site]/[library]/[folder]");

            string sitePath = string.Empty;
            int listStartIndex = 0;

            if (segments[0].Equals("sites", StringComparison.OrdinalIgnoreCase) || 
                segments[0].Equals("teams", StringComparison.OrdinalIgnoreCase))
            {
                sitePath = $"/{segments[0]}/{segments[1]}";
                listStartIndex = 2;
            }
            else
            {
                sitePath = "/";
                listStartIndex = 0;
            }

            // Get the site
            var site = await _graphClient.Sites[$"{hostname}:{sitePath}"].GetAsync();
            if (site == null) throw new Exception("Site not found.");

            // Get all drives (document libraries) in the site
            var drives = await _graphClient.Sites[site.Id].Drives.GetAsync();
            if (drives?.Value == null) throw new Exception("No document libraries found.");

            // Find the drive and folder
            string itemPath = string.Join('/', segments.Skip(listStartIndex));
            itemPath = Uri.UnescapeDataString(itemPath);

            DriveItem? folderItem = null;
            string? driveId = null;

            foreach (var drive in drives.Value)
            {
                try
                {
                    var item = await _graphClient.Drives[drive.Id].Root.ItemWithPath(itemPath).GetAsync();
                    if (item != null)
                    {
                        folderItem = item;
                        driveId = drive.Id;
                        break;
                    }
                }
                catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
                {
                    // Ignore 404, try next drive
                }
            }

            if (folderItem == null || driveId == null)
                throw new Exception("Folder not found in any document library.");

            // Get children of the folder
            var children = await _graphClient.Drives[driveId].Items[folderItem.Id].Children.GetAsync();

            if (children?.Value != null)
            {
                foreach (var child in children.Value)
                {
                    // Only include files, not folders
                    if (child.File != null)
                    {
                        files.Add(new SharePointFile
                        {
                            Name = child.Name ?? "Unknown",
                            Size = child.Size ?? 0,
                            LastModifiedDateTime = child.LastModifiedDateTime
                        });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error retrieving files: {ex.Message}", ex);
        }

        return files;
    }
}
