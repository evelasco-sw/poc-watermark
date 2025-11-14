using Microsoft.AspNetCore.Mvc;
using OfficeIMO.Word;
using WatermarkApi.Utils;

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
                var imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "wwwroot", "images", "watermark.png");

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


    }
}