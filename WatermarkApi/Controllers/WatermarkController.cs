using Microsoft.AspNetCore.Mvc;
using OfficeIMO.Word;
using WatermarkApi.Utils;
using SixLabors.ImageSharp;

namespace WatermarkApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WatermarkController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly ILogger<WatermarkController> _logger;

        public WatermarkController(IWebHostEnvironment webHostEnvironment, ILogger<WatermarkController> logger)
        {
            _webHostEnvironment = webHostEnvironment;
            _logger = logger;
        }

        [HttpPost("word")]
        public async Task<IActionResult> AddWatermarkToWord(IFormFile file, [FromForm] string username)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");
            if (System.IO.Path.GetExtension(file.FileName).ToLower() != ".docx")
                return BadRequest("Only .docx files are supported.");

            try
            {
                var imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "wwwroot", "images", "logo_bcie.png");

                // Guardamos temporalmente el archivo recibido.
                var tempFilePath = Path.GetTempFileName();
                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                using (var document = WordDocument.Load(tempFilePath))
                {
                    document.AddParagraph("Section 0");
                    document.AddHeadersAndFooters();

                    var section0 = document.Sections[0];
                    var section0Header = WatermarkHelper.GetRequiredHeader(section0);

                    section0Header.AddWatermark(WordWatermarkStyle.Image, imagePath);

                    using (var memoryStream = new MemoryStream())
                    {
                        document.SaveAs(memoryStream);
                        return File(memoryStream.ToArray(),
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    "watermarked.docx");
                    }
                }

            }
            catch (Exception ex)
            {
                _logger.LogCritical(ex, "Error al crear la marca de agua {Message}", ex.Message);
                return StatusCode(500, $"Internal error: {ex.Message}");
            }
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

        [HttpGet("png-from-path")]
        public IActionResult GeneratePngFromPath([FromQuery] string path, [FromQuery] string? text = null, [FromQuery] string? fg = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                return BadRequest("Image path is required.");

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
                var bytes = PngTextImageGenerator.CreatePngFromImage(path, text ?? string.Empty, foreground);
                return File(bytes, "image/png", "watermarked.png");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound($"Image not found: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generando PNG desde ruta");
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }


    }
}