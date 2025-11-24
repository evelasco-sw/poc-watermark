using GroupDocs.Watermark;
using GroupDocs.Watermark.Options.Pdf;
using GroupDocs.Watermark.Watermarks;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using Microsoft.VisualBasic;

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

            var loadOptions = new PdfLoadOptions();
            using (Watermarker watermarker = new Watermarker(pdfStream, loadOptions))
            {
                // Add image watermark to the second page
                using (ImageWatermark imageWatermark = new ImageWatermark(imagePath))
                {
                    watermarker.Add(imageWatermark);
                }

                watermarker.Save(outputStream);
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
