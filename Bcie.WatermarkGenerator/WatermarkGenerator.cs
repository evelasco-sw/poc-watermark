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
    /// <summary>
    /// Agrega una marca de agua de imagen a todas las páginas de un documento de Microsoft Word.
    /// </summary>
    /// <param name="wordDocument">Stream del documento de Word de entrada (.doc o .docx). El stream puede no ser seekable; internamente se copia a memoria.</param>
    /// <param name="image">Stream de la imagen a utilizar como marca de agua (formatos comunes soportados por GroupDocs/ImageSharp). El stream puede no ser seekable.</param>
    /// <returns>El documento de Word resultante con la marca de agua aplicada, como arreglo de bytes.</returns>
    /// <exception cref="ArgumentNullException">Se lanza si <paramref name="wordDocument"/> o <paramref name="image"/> es null.</exception>
    /// <remarks>
    /// Este método utiliza GroupDocs.Watermark para insertar la imagen como marca de agua sobre el documento completo.
    /// El contenido de entrada se copia a un <see cref="MemoryStream"/> para garantizar el posicionamiento en el inicio.
    /// </remarks>
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
    /// Genera un PNG a partir de una imagen base, superponiendo la fecha/hora actual y un texto personalizado centrado.
    /// </summary>
    /// <param name="imageStream">Stream de la imagen base (cualquier formato soportado por ImageSharp). Puede no ser seekable.</param>
    /// <param name="customText">Texto personalizado a dibujar. Se coloca debajo de la fecha/hora en líneas separadas.</param>
    /// <returns>Imagen PNG generada como arreglo de bytes.</returns>
    /// <exception cref="ArgumentNullException">Se lanza si <paramref name="imageStream"/> es null.</exception>
    /// <remarks>
    /// El tamaño de la fuente se ajusta automáticamente para que el texto quepa dentro de la imagen con márgenes.
    /// Se intenta usar la fuente Arial; si no está disponible, se usa la primera fuente del sistema.
    /// </remarks>
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
    /// Genera un PNG a partir de una imagen base codificada en Base64, superponiendo la fecha/hora actual y un texto personalizado centrado.
    /// </summary>
    /// <param name="base64Image">Imagen base codificada en Base64.</param>
    /// <param name="customText">Texto personalizado a dibujar. Se coloca debajo de la fecha/hora en líneas separadas.</param>
    /// <returns>PNG generado codificado en Base64.</returns>
    /// <exception cref="ArgumentException">Se lanza si <paramref name="base64Image"/> es null, vacío o contiene solo espacios.</exception>
    public string GeneratePngFromImage(string base64Image, string? customText)
    {
        if (string.IsNullOrWhiteSpace(base64Image)) throw new ArgumentException("Valor requerido", nameof(base64Image));
        using var imgMs = new MemoryStream(Convert.FromBase64String(base64Image));
        var bytes = GeneratePngFromImage(imgMs, customText);
        return Convert.ToBase64String(bytes);
    }

    /// <summary>
    /// Agrega una marca de agua de imagen a un documento de Word provisto en Base64 y devuelve el resultado en Base64.
    /// </summary>
    /// <param name="wordDocumentBase64">Contenido del documento de Word (DOC/DOCX) codificado en Base64.</param>
    /// <param name="imageBase64">Contenido de la imagen de marca de agua codificado en Base64.</param>
    /// <returns>Documento de Word resultante con la marca de agua, codificado en Base64.</returns>
    /// <exception cref="ArgumentException">Se lanza si <paramref name="wordDocumentBase64"/> o <paramref name="imageBase64"/> es null, vacío o contiene solo espacios.</exception>
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
    /// <summary>
    /// Agrega una marca de agua de imagen a todas las diapositivas de una presentación de Microsoft PowerPoint.
    /// </summary>
    /// <param name="presentationDocument">Stream de la presentación de PowerPoint de entrada (.ppt o .pptx). Puede no ser seekable; se copia internamente.</param>
    /// <param name="image">Stream de la imagen a utilizar como marca de agua.</param>
    /// <returns>La presentación resultante con la marca de agua aplicada, como arreglo de bytes.</returns>
    /// <exception cref="ArgumentNullException">Se lanza si <paramref name="presentationDocument"/> o <paramref name="image"/> es null.</exception>
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

    /// <summary>
    /// Agrega una marca de agua de imagen a una presentación de PowerPoint provista en Base64 y devuelve el resultado en Base64.
    /// </summary>
    /// <param name="presentationBase64">Contenido de la presentación (PPT/PPTX) codificado en Base64.</param>
    /// <param name="imageBase64">Contenido de la imagen de marca de agua codificado en Base64.</param>
    /// <returns>Presentación resultante con la marca de agua, codificada en Base64.</returns>
    /// <exception cref="ArgumentException">Se lanza si <paramref name="presentationBase64"/> o <paramref name="imageBase64"/> es null, vacío o contiene solo espacios.</exception>
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
    /// <summary>
    /// Agrega una marca de agua de imagen a todas las páginas de un documento PDF.
    /// </summary>
    /// <param name="pdfDocument">Stream del documento PDF de entrada. Puede no ser seekable; se copia internamente.</param>
    /// <param name="image">Stream de la imagen a utilizar como marca de agua.</param>
    /// <returns>El documento PDF resultante con la marca de agua aplicada, como arreglo de bytes.</returns>
    /// <exception cref="ArgumentNullException">Se lanza si <paramref name="pdfDocument"/> o <paramref name="image"/> es null.</exception>
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

    /// <summary>
    /// Agrega una marca de agua de imagen a un documento PDF provisto en Base64 y devuelve el resultado en Base64.
    /// </summary>
    /// <param name="pdfBase64">Contenido del documento PDF codificado en Base64.</param>
    /// <param name="imageBase64">Contenido de la imagen de marca de agua codificado en Base64.</param>
    /// <returns>Documento PDF resultante con la marca de agua, codificado en Base64.</returns>
    /// <exception cref="ArgumentException">Se lanza si <paramref name="pdfBase64"/> o <paramref name="imageBase64"/> es null, vacío o contiene solo espacios.</exception>
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
