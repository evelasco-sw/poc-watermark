using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace WatermarkApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WatermarkController : ControllerBase
    {

        [HttpPost("word")]
        public async Task<IActionResult> AddWatermarkToWord(IFormFile file, [FromForm] string username)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");
            if (System.IO.Path.GetExtension(file.FileName).ToLower() != ".docx")
                return BadRequest("Only .docx files are supported.");

            try
            {
                // 1. Generar imagen de marca de agua dinámica
                var watermarkImage = GenerateWatermarkImage(username, DateTime.UtcNow);

                // 2. Procesar el documento con Open XML SDK
                using var memoryStream = new MemoryStream();
                await file.CopyToAsync(memoryStream);
                memoryStream.Position = 0;

                AddWatermarkToDocument(memoryStream, watermarkImage);

                // 3. Devolver el documento modificado
                memoryStream.Position = 0;
                return File(memoryStream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            "watermarked.docx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal error: {ex.Message}");
            }
        }

        private byte[] GenerateWatermarkImage(string username, DateTime timestamp)
        {
            // Tamaño óptimo para una marca de agua (se ajustará en Word)
            var width = 400;
            var height = 200;

            using var image = new Image<Rgba32>(width, height);
           
            using var ms = new MemoryStream();
            image.SaveAsPng(ms);
            return ms.ToArray();
        }

        private void AddWatermarkToDocument(Stream documentStream, byte[] watermarkImage)
        {
            using var doc = WordprocessingDocument.Open(documentStream, true);
            var mainPart = doc.MainDocumentPart;

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new MemoryStream(watermarkImage))
            {
                imagePart.FeedData(imageStream);
            }
            var imageId = mainPart.GetIdOfPart(imagePart);

            // 3. Aplicar la marca de agua a TODOS los tipos de header
            ApplyWatermarkToAllHeaders(mainPart, imageId);
        }

        private void ApplyWatermarkToAllHeaders(MainDocumentPart mainPart, string imageId)
        {
            // Tipos de header: Default, First, Even
            foreach (var headerPart in mainPart.HeaderParts)
            {
                var header = headerPart.Header;
                header.RemoveAllChildren(); // Limpiar contenido existente (solo para el ejemplo)

                // Crear un párrafo para la marca de agua
                var paragraph = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "Header" },
                        new SpacingBetweenLines { After = "0" },
                        new Justification { Val = JustificationValues.Center }
                    )
                );

                // Crear el contenedor VML para la imagen (compatible con Word)
                var shape = new DocumentFormat.OpenXml.Vml.Shape
                {
                    Id = "Watermark_" + Guid.NewGuid().ToString("N"),
                    Style = "position:absolute;margin-left:0;margin-top:0;width:100%;height:100%;z-index:-100",
                    WrapCoordinates = "left 0,top 0,right 0,bottom 0",
                    FillColor = "transparent",
                    StrokeWeight = "0",
                    // Rotation = 30, // Rotación diagonal
                    AllowOverlap = true,
                    // BehindDocument = true // ¡Clave para que esté detrás del texto!
                };

                // Enlazar la imagen al shape
                shape.Append(new DocumentFormat.OpenXml.Vml.ImageData
                {
                    Title = "Watermark",
                    Id = imageId
                });

                // Añadir el shape al párrafo
                paragraph.Append(new Run(shape));
                header.Append(paragraph);
                header.Save();
            }
        }
    }
}