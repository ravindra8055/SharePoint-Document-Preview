namespace SharePointViewer.Models;

public class SharePointFile
{
    public string Name { get; set; } = string.Empty;
    public long Size { get; set; }
    public DateTimeOffset? LastModifiedDateTime { get; set; }
    public string? PreviewUrl { get; set; }
}
