using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using PnP.Core.Auth;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using SharePointViewer.Models;
using System.Security;

namespace SharePointViewer.Services;

public class SharePointService
{
    private readonly IPnPContextFactory _pnpContextFactory;
    private readonly IConfiguration _configuration;

    public SharePointService(IConfiguration configuration, IPnPContextFactory pnpContextFactory)
    {
        _configuration = configuration;
        _pnpContextFactory = pnpContextFactory;
    }

    /// <summary>
    /// Infers the Tenant ID (e.g., contoso.onmicrosoft.com) from the SharePoint URL (e.g., contoso.sharepoint.com).
    /// This mimics how PnP PowerShell connects without needing an explicit Tenant ID.
    /// </summary>
    private string GetTenantIdFromUrl(string url)
    {
        var configuredTenant = _configuration["SharePointAuth:TenantId"];
        if (!string.IsNullOrWhiteSpace(configuredTenant) && configuredTenant != "YOUR_TENANT_ID")
        {
            return configuredTenant;
        }

        if (string.IsNullOrWhiteSpace(url)) return "organizations";

        try
        {
            var uri = new Uri(url);
            if (uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase))
            {
                return uri.Host.Replace(".sharepoint.com", ".onmicrosoft.com", StringComparison.OrdinalIgnoreCase);
            }
        }
        catch
        {
            // Ignore invalid URI, fallback to organizations
        }
        
        return "organizations"; // Fallback for MSAL
    }

    private GraphServiceClient GetGraphClient(string url)
    {
        var authType = _configuration["SharePointAuth:AuthType"];
        var tenantId = GetTenantIdFromUrl(url);
        var clientId = _configuration["SharePointAuth:ClientId"];
        var scopes = new[] { "https://graph.microsoft.com/.default" };

        if (string.Equals(authType, "UsernamePassword", StringComparison.OrdinalIgnoreCase))
        {
            var username = _configuration["SharePointAuth:Username"];
            var password = _configuration["SharePointAuth:Password"];

            var options = new UsernamePasswordCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
            var cred = new UsernamePasswordCredential(username, password, tenantId, clientId, options);
            return new GraphServiceClient(cred, scopes);
        }
        else
        {
            var clientSecret = _configuration["SharePointAuth:ClientSecret"];
            var options = new ClientSecretCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
            var cred = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
            return new GraphServiceClient(cred, scopes);
        }
    }

    private IAuthenticationProvider GetPnPAuthenticationProvider(string url)
    {
        var authType = _configuration["SharePointAuth:AuthType"];
        var tenantId = GetTenantIdFromUrl(url);
        var clientId = _configuration["SharePointAuth:ClientId"];

        if (string.Equals(authType, "UsernamePassword", StringComparison.OrdinalIgnoreCase))
        {
            var username = _configuration["SharePointAuth:Username"];
            var password = _configuration["SharePointAuth:Password"];

            var securePassword = new SecureString();
            if (!string.IsNullOrEmpty(password))
            {
                foreach (char c in password) securePassword.AppendChar(c);
            }

            return new UsernamePasswordAuthenticationProvider(clientId, tenantId, username, securePassword);
        }
        else
        {
            var clientSecret = _configuration["SharePointAuth:ClientSecret"];
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            return new ExternalAuthenticationProvider((resource, scopes) => 
            {
                var tokenRequestContext = new Azure.Core.TokenRequestContext(scopes);
                var token = credential.GetToken(tokenRequestContext);
                return Task.FromResult(token.Token);
            });
        }
    }

    public async Task<List<SharePointFile>> GetFilesInFolderAsync(string folderUrl)
    {
        var files = new List<SharePointFile>();

        try
        {
            var uri = new Uri(folderUrl);
            var hostname = uri.Host;
            var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            
            if (segments.Length < 3)
                throw new ArgumentException("Invalid SharePoint folder URL format.");

            string sitePath = (segments[0].Equals("sites", StringComparison.OrdinalIgnoreCase) || segments[0].Equals("teams", StringComparison.OrdinalIgnoreCase))
                ? $"/{segments[0]}/{segments[1]}"
                : "/";

            string siteUrl = $"https://{hostname}{sitePath}";
            string folderRelativeUrl = uri.AbsolutePath;

            // 1. Create PnP Context using the dynamically resolved Tenant ID
            var authProvider = GetPnPAuthenticationProvider(folderUrl);
            using var context = await _pnpContextFactory.CreateAsync(new Uri(siteUrl), authProvider);

            // 2. Get the folder and its files using PnP Core
            var folder = await context.Web.GetFolderByServerRelativeUrlAsync(folderRelativeUrl, 
                f => f.Files.QueryProperties(
                    file => file.UniqueId,
                    file => file.Name,
                    file => file.Length,
                    file => file.TimeLastModified,
                    file => file.ServerRelativeUrl
                ));

            foreach (var file in folder.Files)
            {
                files.Add(new SharePointFile
                {
                    Id = file.UniqueId.ToString(),
                    DriveId = "graph-shares-api", // Dummy value to pass UI checks
                    Name = file.Name,
                    Size = file.Length,
                    LastModifiedDateTime = file.TimeLastModified,
                    PreviewUrl = $"https://{hostname}{file.ServerRelativeUrl}" // Absolute URL for Graph Shares API
                });
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error retrieving files using PnP Core: {ex.Message}", ex);
        }

        return files;
    }

    public async Task<string?> GetFilePreviewUrlAsync(string driveId, string itemId, string siteUrl)
    {
        try
        {
            var graphClient = GetGraphClient(siteUrl);
            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Preview.PreviewPostRequestBody();
            var result = await graphClient.Drives[driveId].Items[itemId].Preview.PostAsync(requestBody);
            return result?.GetUrl;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting preview URL from Graph API: {ex.Message}");
            return null;
        }
    }

    private string EncodeSharePointUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url)) return string.Empty;
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(url);
        string base64 = Convert.ToBase64String(bytes);
        return "u!" + base64.TrimEnd('=').Replace('/', '_').Replace('+', '-');
    }

    public async Task<string?> GetPreviewUrlFromSharePointUrlAsync(string sharePointUrl)
    {
        try
        {
            var graphClient = GetGraphClient(sharePointUrl);
            string encodedUrl = EncodeSharePointUrl(sharePointUrl);
            var driveItem = await graphClient.Shares[encodedUrl].DriveItem.GetAsync();

            if (driveItem?.ParentReference?.DriveId == null || driveItem?.Id == null)
                return null;

            var requestBody = new Microsoft.Graph.Drives.Item.Items.Item.Preview.PreviewPostRequestBody();
            var result = await graphClient.Drives[driveItem.ParentReference.DriveId]
                                           .Items[driveItem.Id]
                                           .Preview
                                           .PostAsync(requestBody);
            return result?.GetUrl;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            return null;
        }
        catch (Exception ex)
        {
            return null;
        }
    }
}
