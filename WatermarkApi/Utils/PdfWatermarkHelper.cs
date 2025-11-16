using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Geom;
using iText.IO.Image;
using iText.Layout;
using iText.Layout.Element;

namespace WatermarkApi.Utils;

public static class PdfWatermarkHelper
{
    /// <summary>
    /// Adds an image watermark to all pages in a PDF document from a file path.
    /// </summary>
    public static byte[] AddWatermarkToPdf(string pdfPath, string imagePath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath))
            throw new ArgumentException("PDF path cannot be null or empty.", nameof(pdfPath));
        if (!System.IO.File.Exists(pdfPath))
            throw new FileNotFoundException($"PDF file not found: {pdfPath}");
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var fileStream = new FileStream(pdfPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        return AddWatermarkToPdf(fileStream, imagePath);
    }

    /// <summary>
    /// Adds an image watermark to all pages in a PDF document from a stream.
    /// </summary>
    public static byte[] AddWatermarkToPdf(Stream pdfStream, string imagePath)
    {
        if (pdfStream == null)
            throw new ArgumentNullException(nameof(pdfStream));
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var memoryStream = new MemoryStream();
        var outputStream = new MemoryStream();

        try
        {
            pdfStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            using (PdfReader reader = new PdfReader(memoryStream))
            using (PdfWriter writer = new PdfWriter(outputStream))
            using (PdfDocument pdfDoc = new PdfDocument(reader, writer))
            {
                // Leer la imagen
                ImageData imageData = ImageDataFactory.Create(imagePath);

                // Procesar cada página
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    PdfPage page = pdfDoc.GetPage(i);
                    Rectangle pageSize = page.GetPageSize();
                    
                    // Crear un canvas para dibujar sobre la página
                    PdfCanvas canvas = new PdfCanvas(page);
                    
                    // Calcular posición y tamaño de la imagen
                    float pageWidth = pageSize.GetWidth();
                    float pageHeight = pageSize.GetHeight();
                    
                    float imageWidth = Math.Min(imageData.GetWidth(), pageWidth * 0.8f);
                    float imageHeight = (imageWidth / imageData.GetWidth()) * imageData.GetHeight();
                    
                    // Si la altura excede el 80% de la página, escalar por altura
                    if (imageHeight > pageHeight * 0.8f)
                    {
                        imageHeight = pageHeight * 0.8f;
                        imageWidth = (imageHeight / imageData.GetHeight()) * imageData.GetWidth();
                    }
                    
                    float x = (pageWidth - imageWidth) / 2f;
                    float y = (pageHeight - imageHeight) / 2f;
                    
                    // Crear imagen y agregarla con Document API
                    Image image = new Image(imageData);
                    image.SetWidth(imageWidth);
                    image.SetHeight(imageHeight);
                    image.SetFixedPosition(x, y);
                    
                    // Usar Document para agregar la imagen
                    Document doc = new Document(pdfDoc);
                    doc.Add(image);
                    doc.Close();
                }
            }

            return outputStream.ToArray();
        }
        finally
        {
            memoryStream?.Dispose();
            outputStream?.Dispose();
        }
    }
}
