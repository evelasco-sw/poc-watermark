using GroupDocs.Watermark;
using GroupDocs.Watermark.Options.WordProcessing;
using GroupDocs.Watermark.Watermarks;

namespace WatermarkApi.Utils;

public static class WordWatermarkHelper
{
    /// <summary>
    /// Agrega una marca de agua de imagen a todas las páginas de un documento de Word desde una ruta de archivo.
    /// </summary>
    public static byte[] AddWatermarkToWord(string documentPath, string imagePath)
    {
        if (string.IsNullOrWhiteSpace(documentPath))
            throw new ArgumentException("Document path cannot be null or empty.", nameof(documentPath));
        if (!System.IO.File.Exists(documentPath))
            throw new FileNotFoundException($"Word file not found: {documentPath}");
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var fileStream = new FileStream(documentPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        return AddWatermarkToWord(fileStream, imagePath);
    }

    /// <summary>
    /// Agrega una marca de agua de imagen a todas las páginas de un documento de Word desde un Stream.
    /// </summary>
    public static byte[] AddWatermarkToWord(Stream documentStream, string imagePath)
    {
        if (documentStream == null)
            throw new ArgumentNullException(nameof(documentStream));
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var inputCopy = new MemoryStream();
        var outputStream = new MemoryStream();

        try
        {
            // Copiar el documento al MemoryStream para asegurar posicionamiento
            documentStream.CopyTo(inputCopy);
            inputCopy.Position = 0;

            var loadOptions = new WordProcessingLoadOptions();
            using (var watermarker = new Watermarker(inputCopy, loadOptions))
            {
                using (var imageWatermark = new ImageWatermark(imagePath))
                {
                    watermarker.Add(imageWatermark);
                }

                watermarker.Save(outputStream);
            }

            return outputStream.ToArray();
        }
        finally
        {
            inputCopy?.Dispose();
            outputStream?.Dispose();
        }
    }
}