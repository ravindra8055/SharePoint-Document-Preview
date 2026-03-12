namespace SharePointViewer.Models;

public class SharePointFile
{
    public string? Id { get; set; }
    public string? DriveId { get; set; }
    public string Name { get; set; } = string.Empty;
    public long Size { get; set; }
    public DateTimeOffset? LastModifiedDateTime { get; set; }
    public string? PreviewUrl { get; set; }
    public string? EmbedUrl { get; set; }
    public string? WebUrl { get; set; }
    public string? DocUrl { get; set; }
}
