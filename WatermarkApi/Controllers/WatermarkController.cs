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
        public async Task<IActionResult> AddWatermarkToWord(IFormFile docFile, IFormFile? imageFile = null)
        {
            if (docFile == null || docFile.Length == 0)
                return BadRequest("No Word document uploaded.");
            if (System.IO.Path.GetExtension(docFile.FileName).ToLower() != ".docx")
                return BadRequest("Only .docx files are supported.");

            string? imagePath = null;
            string? tempImagePath = null;

            try
            {
                // Si se proporciona una imagen personalizada, guardarla temporalmente
                if (imageFile != null && imageFile.Length > 0)
                {
                    tempImagePath = Path.GetTempFileName();
                    using (var stream = new FileStream(tempImagePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }
                    imagePath = tempImagePath;
                }
                else
                {
                    // Usar imagen por defecto si no se proporciona una personalizada
                    imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "wwwroot", "images", "logo_bcie.png");
                }

                // Guardar temporalmente el documento recibido
                var tempDocPath = Path.GetTempFileName();
                using (var stream = new FileStream(tempDocPath, FileMode.Create))
                {
                    await docFile.CopyToAsync(stream);
                }

                using (var document = WordDocument.Load(tempDocPath))
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
            finally
            {
                // Limpiar archivos temporales
                if (!string.IsNullOrEmpty(tempImagePath) && System.IO.File.Exists(tempImagePath))
                {
                    try { System.IO.File.Delete(tempImagePath); } catch { }
                }
            }
        }
        
        [HttpPost("powerpoint")]
        public async Task<IActionResult> AddWatermarkToPowerPoint(IFormFile pptFile, IFormFile? imageFile = null)
        {
            if (pptFile == null || pptFile.Length == 0)
                return BadRequest("No PowerPoint presentation uploaded.");
            
            string ext = System.IO.Path.GetExtension(pptFile.FileName).ToLower();
            if (ext != ".pptx" && ext != ".ppt")
                return BadRequest("Only .pptx or .ppt files are supported.");

            string? imagePath = null;
            string? tempImagePath = null;
            string? tempPptPath = null;

            try
            {
                // Si se proporciona una imagen personalizada, guardarla temporalmente
                if (imageFile != null && imageFile.Length > 0)
                {
                    tempImagePath = Path.GetTempFileName();
                    using (var stream = new FileStream(tempImagePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }
                    imagePath = tempImagePath;
                }
                else
                {
                    // Usar imagen por defecto si no se proporciona una personalizada
                    imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "wwwroot", "images", "logo_bcie.png");
                }

                // Guardar temporalmente la presentación recibida
                tempPptPath = Path.GetTempFileName();
                using (var stream = new FileStream(tempPptPath, FileMode.Create))
                {
                    await pptFile.CopyToAsync(stream);
                }

                // Agregar marca de agua
                var bytes = PowerPointWatermarkHelper.AddWatermarkToPresentation(tempPptPath, imagePath);
                return File(bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "watermarked.pptx");
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogError(ex, "Archivo no encontrado: {Message}", ex.Message);
                return NotFound($"File not found: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogCritical(ex, "Error al agregar marca de agua a PowerPoint: {Message}", ex.Message);
                return StatusCode(500, $"Internal error: {ex.Message}");
            }
            finally
            {
                // Limpiar archivos temporales
                if (!string.IsNullOrEmpty(tempImagePath) && System.IO.File.Exists(tempImagePath))
                {
                    try { System.IO.File.Delete(tempImagePath); } catch { }
                }
                if (!string.IsNullOrEmpty(tempPptPath) && System.IO.File.Exists(tempPptPath))
                {
                    try { System.IO.File.Delete(tempPptPath); } catch { }
                }
            }
        }

        [HttpPost("pdf")]
        public async Task<IActionResult> AddWatermarkToPdf(IFormFile pdfFile, IFormFile? imageFile = null)
        {
            if (pdfFile == null || pdfFile.Length == 0)
                return BadRequest("No PDF file uploaded.");
            
            if (System.IO.Path.GetExtension(pdfFile.FileName).ToLower() != ".pdf")
                return BadRequest("Only .pdf files are supported.");

            string? imagePath = null;
            string? tempImagePath = null;
            string? tempPdfPath = null;

            try
            {
                // Si se proporciona una imagen personalizada, guardarla temporalmente
                if (imageFile != null && imageFile.Length > 0)
                {
                    tempImagePath = Path.GetTempFileName();
                    using (var stream = new FileStream(tempImagePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }
                    imagePath = tempImagePath;
                }
                else
                {
                    // Usar imagen por defecto si no se proporciona una personalizada
                    imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "wwwroot", "images", "logo_bcie.png");
                }

                // Guardar temporalmente el PDF recibido
                tempPdfPath = Path.GetTempFileName();
                using (var stream = new FileStream(tempPdfPath, FileMode.Create))
                {
                    await pdfFile.CopyToAsync(stream);
                }

                // Agregar marca de agua
                var bytes = PdfWatermarkHelper.AddWatermarkToPdf(tempPdfPath, imagePath);
                return File(bytes, "application/pdf", "watermarked.pdf");
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogError(ex, "Archivo no encontrado: {Message}", ex.Message);
                return NotFound($"File not found: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogCritical(ex, "Error al agregar marca de agua a PDF: {Message}", ex.Message);
                return StatusCode(500, $"Internal error: {ex.Message}");
            }
            finally
            {
                // Limpiar archivos temporales
                if (!string.IsNullOrEmpty(tempImagePath) && System.IO.File.Exists(tempImagePath))
                {
                    try { System.IO.File.Delete(tempImagePath); } catch { }
                }
                if (!string.IsNullOrEmpty(tempPdfPath) && System.IO.File.Exists(tempPdfPath))
                {
                    try { System.IO.File.Delete(tempPdfPath); } catch { }
                }
            }
        }

    }
}