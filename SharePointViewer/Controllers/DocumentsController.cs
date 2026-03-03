using Microsoft.AspNetCore.Mvc;
using SharePointViewer.Services;

namespace SharePointViewer.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DocumentsController : ControllerBase
{
    private readonly SharePointService _sharePointService;

    public DocumentsController(SharePointService sharePointService)
    {
        _sharePointService = sharePointService;
    }

    public class PreviewRequest
    {
        public string Url { get; set; } = string.Empty;
    }

    [HttpPost("preview")]
    public async Task<IActionResult> GetPreview([FromBody] PreviewRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Url))
        {
            return BadRequest(new { error = "URL is required." });
        }

        var previewUrl = await _sharePointService.GetPreviewUrlFromSharePointUrlAsync(request.Url);

        if (string.IsNullOrEmpty(previewUrl))
        {
            return NotFound(new { error = "Could not generate preview URL. The file may not exist or permissions are insufficient." });
        }

        return Ok(new { previewUrl });
    }
}
