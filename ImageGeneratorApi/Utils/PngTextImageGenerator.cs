using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.Processing;

namespace ImageGeneratorApi.Utils;

public static class PngTextImageGenerator
{
    /// <summary>
    /// Creates a new PNG image with specified dimensions and overlays text.
    /// </summary>
    public static byte[] CreateSquarePng(int size, string customText, Color? backgroundColor = null, Color? textColor = null)
    {
        if (size <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(size));
        }

        customText ??= string.Empty;
        var dateLine = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var combined = dateLine + "\n" + customText.Trim();

        var margin = Math.Max(20, size / 20);
        float fontSize = size / 8f;
        if (fontSize < 8f) fontSize = 8f;

        FontFamily fontFamily;
        if (!SystemFonts.TryGet("Arial", out fontFamily))
        {
            var fallbackFamily = SystemFonts.Families.FirstOrDefault();
            fontFamily = fallbackFamily;
        }

        var font = fontFamily.CreateFont(fontSize);

        RichTextOptions options = new(font)
        {
            WrappingLength = size - margin * 2,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top,
        };

        var attempts = 0;
        const int maxAttempts = 50;
        var measure = TextMeasurer.MeasureSize(combined, options);

        while ((measure.Width > size - margin * 2 || measure.Height > size - margin * 2) && attempts < maxAttempts)
        {
            fontSize *= 0.92f;
            if (fontSize < 8f) break;
            font = fontFamily.CreateFont(fontSize);
            options = new RichTextOptions(font)
            {
                WrappingLength = size - margin * 2,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
            };
            measure = TextMeasurer.MeasureSize(combined, options);
            attempts++;
        }

        using var image = new Image<Rgba32>(size, size);
        image.Mutate(ctx => ctx.Fill(backgroundColor ?? Color.White));

        var startY = (size - measure.Height) / 2f;
        options.Origin = new PointF(size / 2f, startY);

        image.Mutate(ctx => ctx.DrawText(options, combined, textColor ?? Color.Black));

        using var ms = new MemoryStream();
        image.Save(ms, new PngEncoder());
        return ms.ToArray();
    }

    /// <summary>
    /// Loads an existing image from file path and overlays text on it.
    /// </summary>
    public static byte[] CreatePngFromImage(string imagePath, string customText, Color? textColor = null)
    {
        if (string.IsNullOrWhiteSpace(imagePath))
        {
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        }

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"Image file not found: {imagePath}");
        }

        return CreatePngFromImageStream(new FileStream(imagePath, FileMode.Open, FileAccess.Read), customText, textColor, disposeStream: true);
    }

    /// <summary>
    /// Loads an image from a stream and overlays text on it.
    /// </summary>
    public static byte[] CreatePngFromImageStream(Stream imageStream, string customText, Color? textColor = null, bool disposeStream = false)
    {
        if (imageStream == null)
        {
            throw new ArgumentNullException(nameof(imageStream));
        }

        customText ??= string.Empty;
        var dateLine = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var combined = dateLine + "\n" + customText.Trim();

        try
        {
            using var image = Image.Load<Rgba32>(imageStream);
            int imageSize = Math.Min(image.Width, image.Height);
            var margin = Math.Max(20, imageSize / 20);
            float fontSize = imageSize / 8f;
            if (fontSize < 8f) fontSize = 8f;

            FontFamily fontFamily;
            if (!SystemFonts.TryGet("Arial", out fontFamily))
            {
                var fallbackFamily = SystemFonts.Families.FirstOrDefault();
                fontFamily = fallbackFamily;
            }

            var font = fontFamily.CreateFont(fontSize);

            RichTextOptions options = new(font)
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

            image.Mutate(ctx => ctx.DrawText(options, combined, textColor ?? Color.Black));

            using var ms = new MemoryStream();
            image.Save(ms, new PngEncoder());
            return ms.ToArray();
        }
        finally
        {
            if (disposeStream)
            {
                imageStream.Dispose();
            }
        }
    }
}