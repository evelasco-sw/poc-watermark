using System;
using System.IO;
using System.Linq;
using GroupDocs.Watermark;
using GroupDocs.Watermark.Options.Pdf;
using GroupDocs.Watermark.Options.Presentation;
using GroupDocs.Watermark.Options.WordProcessing;
using GroupDocs.Watermark.Watermarks;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing.Processing;

namespace Bcie.WatermarkGenerator;

/// <summary>
/// Servicio para aplicar marcas de agua de imagen a documentos Word, PowerPoint y PDF.
/// Provee métodos que aceptan entradas como Stream y como cadenas Base64.
/// </summary>
public class WatermarkGenerator
{
    // -------- Word --------
    public byte[] AddWatermarkToWord(Stream wordDocument, Stream image)
    {
        if (wordDocument is null) throw new ArgumentNullException(nameof(wordDocument));
        if (image is null) throw new ArgumentNullException(nameof(image));

        using var inputCopy = ToReadableMemoryStream(wordDocument);
        var output = new MemoryStream();

        var loadOptions = new WordProcessingLoadOptions();
        using (var watermarker = new Watermarker(inputCopy, loadOptions))
        using (var imageWatermark = new ImageWatermark(ToReadableMemoryStream(image)))
        {
            watermarker.Add(imageWatermark);
            watermarker.Save(output);
        }

        return output.ToArray();
    }

    // -------- Generación de imágenes (similar a ImageGeneratorApi) --------
    /// <summary>
    /// Genera un PNG a partir de una imagen base, superponiendo fecha/hora y texto personalizado centrado.
    /// </summary>
    /// <param name="imageStream">Stream de la imagen base (cualquier formato soportado por ImageSharp).</param>
    /// <param name="customText">Texto personalizado a dibujar. Se concatena con la fecha en la primera línea.</param>
    /// <returns>PNG como arreglo de bytes.</returns>
    public byte[] GeneratePngFromImage(Stream imageStream, string? customText)
    {
        if (imageStream is null) throw new ArgumentNullException(nameof(imageStream));

        customText ??= string.Empty;
        var dateLine = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var combined = dateLine + "\n" + customText.Trim();

        using var readable = ToReadableMemoryStream(imageStream);
        using var image = Image.Load<Rgba32>(readable);

        int imageSize = Math.Min(image.Width, image.Height);
        var margin = Math.Max(20, imageSize / 20);
        float fontSize = imageSize / 8f;
        if (fontSize < 8f) fontSize = 8f;

        if (!SystemFonts.TryGet("Arial", out FontFamily fontFamily))
        {
            fontFamily = SystemFonts.Families.FirstOrDefault();
        }

        var font = fontFamily.CreateFont(fontSize);
        var options = new RichTextOptions(font)
        {
            WrappingLength = imageSize - margin * 2,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top,
        };

        var attempts = 0;
        const int maxAttempts = 50;
        var measure = TextMeasurer.MeasureSize(combined, options);
        while ((measure.Width > imageSize - margin * 2 || measure.Height > imageSize - margin * 2) && attempts < maxAttempts)
        {
            fontSize *= 0.92f;
            if (fontSize < 8f) break;
            font = fontFamily.CreateFont(fontSize);
            options = new RichTextOptions(font)
            {
                WrappingLength = imageSize - margin * 2,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
            };
            measure = TextMeasurer.MeasureSize(combined, options);
            attempts++;
        }

        var startY = (imageSize - measure.Height) / 2f;
        options.Origin = new PointF(imageSize / 2f, startY);

        image.Mutate(ctx => ctx.DrawText(options, combined, SixLabors.ImageSharp.Color.Black));

        using var ms = new MemoryStream();
        image.Save(ms, new PngEncoder());
        return ms.ToArray();
    }

    /// <summary>
    /// Genera un PNG a partir de una imagen base codificada en Base64, superponiendo fecha/hora y texto personalizado centrado.
    /// </summary>
    /// <param name="base64Image">Imagen base en Base64.</param>
    /// <param name="customText">Texto personalizado a dibujar.</param>
    /// <returns>PNG en Base64.</returns>
    public string GeneratePngFromImage(string base64Image, string? customText)
    {
        if (string.IsNullOrWhiteSpace(base64Image)) throw new ArgumentException("Valor requerido", nameof(base64Image));
        using var imgMs = new MemoryStream(Convert.FromBase64String(base64Image));
        var bytes = GeneratePngFromImage(imgMs, customText);
        return Convert.ToBase64String(bytes);
    }

    public string AddWatermarkToWord(string wordDocumentBase64, string imageBase64)
    {
        if (string.IsNullOrWhiteSpace(wordDocumentBase64)) throw new ArgumentException("Valor requerido", nameof(wordDocumentBase64));
        if (string.IsNullOrWhiteSpace(imageBase64)) throw new ArgumentException("Valor requerido", nameof(imageBase64));

        using var docMs = new MemoryStream(Convert.FromBase64String(wordDocumentBase64));
        using var imgMs = new MemoryStream(Convert.FromBase64String(imageBase64));
        var result = AddWatermarkToWord(docMs, imgMs);
        return Convert.ToBase64String(result);
    }

    // -------- PowerPoint --------
    public byte[] AddWatermarkToPresentation(Stream presentationDocument, Stream image)
    {
        if (presentationDocument is null) throw new ArgumentNullException(nameof(presentationDocument));
        if (image is null) throw new ArgumentNullException(nameof(image));

        using var inputCopy = ToReadableMemoryStream(presentationDocument);
        var output = new MemoryStream();

        var loadOptions = new PresentationLoadOptions();
        using (var watermarker = new Watermarker(inputCopy, loadOptions))
        using (var imageWatermark = new ImageWatermark(ToReadableMemoryStream(image)))
        {
            watermarker.Add(imageWatermark);
            watermarker.Save(output);
        }

        return output.ToArray();
    }

    public string AddWatermarkToPresentation(string presentationBase64, string imageBase64)
    {
        if (string.IsNullOrWhiteSpace(presentationBase64)) throw new ArgumentException("Valor requerido", nameof(presentationBase64));
        if (string.IsNullOrWhiteSpace(imageBase64)) throw new ArgumentException("Valor requerido", nameof(imageBase64));

        using var docMs = new MemoryStream(Convert.FromBase64String(presentationBase64));
        using var imgMs = new MemoryStream(Convert.FromBase64String(imageBase64));
        var result = AddWatermarkToPresentation(docMs, imgMs);
        return Convert.ToBase64String(result);
    }

    // -------- PDF --------
    public byte[] AddWatermarkToPdf(Stream pdfDocument, Stream image)
    {
        if (pdfDocument is null) throw new ArgumentNullException(nameof(pdfDocument));
        if (image is null) throw new ArgumentNullException(nameof(image));

        using var inputCopy = ToReadableMemoryStream(pdfDocument);
        var output = new MemoryStream();

        var loadOptions = new PdfLoadOptions();
        using (var watermarker = new Watermarker(inputCopy, loadOptions))
        using (var imageWatermark = new ImageWatermark(ToReadableMemoryStream(image)))
        {
            watermarker.Add(imageWatermark);
            watermarker.Save(output);
        }

        return output.ToArray();
    }

    public string AddWatermarkToPdf(string pdfBase64, string imageBase64)
    {
        if (string.IsNullOrWhiteSpace(pdfBase64)) throw new ArgumentException("Valor requerido", nameof(pdfBase64));
        if (string.IsNullOrWhiteSpace(imageBase64)) throw new ArgumentException("Valor requerido", nameof(imageBase64));

        using var docMs = new MemoryStream(Convert.FromBase64String(pdfBase64));
        using var imgMs = new MemoryStream(Convert.FromBase64String(imageBase64));
        var result = AddWatermarkToPdf(docMs, imgMs);
        return Convert.ToBase64String(result);
    }

    // -------- Helpers --------
    private static MemoryStream ToReadableMemoryStream(Stream source)
    {
        if (source is MemoryStream ms && ms.CanSeek)
        {
            // Ensure position at start
            if (ms.Position != 0) ms.Position = 0;
            return ms;
        }

        var copy = new MemoryStream();
        source.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }

    private static byte[] ReadAllBytes(Stream source)
    {
        if (source is MemoryStream ms && ms.TryGetBuffer(out var seg))
        {
            return seg.ToArray();
        }
        using var tmp = new MemoryStream();
        source.CopyTo(tmp);
        return tmp.ToArray();
    }
}
