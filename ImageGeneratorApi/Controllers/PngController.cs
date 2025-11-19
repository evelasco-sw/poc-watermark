using ImageGeneratorApi.Utils;
using Microsoft.AspNetCore.Mvc;
using SixLabors.ImageSharp;

namespace ImageGeneratorApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class PngController : ControllerBase
{
    private readonly ILogger<PngController> _logger;

    public PngController(ILogger<PngController> logger)
    {
        _logger = logger;
    }
    
    [HttpGet("png")]
        public IActionResult GeneratePng([FromQuery] int size = 512, [FromQuery] string? text = null, [FromQuery] string? bg = null, [FromQuery] string? fg = null)
        {
            if (size <= 0 || size > 5000)
            {
                return BadRequest("Size debe estar entre 1 y 5000.");
            }

            Color background;
            Color foreground;
            try
            {
                background = string.IsNullOrWhiteSpace(bg) ? Color.White : Color.ParseHex(bg.Trim());
                foreground = string.IsNullOrWhiteSpace(fg) ? Color.Black : Color.ParseHex(fg.Trim());
            }
            catch (Exception ex)
            {
                return BadRequest($"Color inválido: {ex.Message}");
            }

            try
            {
                var bytes = PngTextImageGenerator.CreateSquarePng(size, text ?? string.Empty, background, foreground);
                return File(bytes, "image/png", "generated.png");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generando PNG");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        [HttpPost("png-from-file")]
        public async Task<IActionResult> GeneratePngFromFile(IFormFile file, [FromForm] string? text = null, [FromForm] string? fg = null)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No image file uploaded.");

            Color foreground;
            try
            {
                foreground = string.IsNullOrWhiteSpace(fg) ? Color.Black : Color.ParseHex(fg.Trim());
            }
            catch (Exception ex)
            {
                return BadRequest($"Color inválido: {ex.Message}");
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;
                var bytes = PngTextImageGenerator.CreatePngFromImageStream(stream, text ?? string.Empty, foreground, disposeStream: false);
                return File(bytes, "image/png", "watermarked.png");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generando PNG desde archivo");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }
}