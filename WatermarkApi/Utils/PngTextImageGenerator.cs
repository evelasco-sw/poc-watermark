using System;
using System.IO;
using System.Linq;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.Processing;

namespace WatermarkApi.Utils;

public static class PngTextImageGenerator
{
    public static byte[] CreateSquarePng(int size, string customText, Color? backgroundColor = null, Color? textColor = null)
    {
        if (size <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(size));
        }

        customText ??= string.Empty;

        var dateLine = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        var fontFamily = SystemFonts.Collection.Families
            .FirstOrDefault(f => f.Name.Equals("Arial", StringComparison.OrdinalIgnoreCase))
            ?? SystemFonts.Collection.Families.First();

        var margin = Math.Max(20, size / 20);
        float fontSize = size / 8f;
        if (fontSize < 8f) fontSize = 8f;

        Font font = fontFamily.CreateFont(fontSize);

        var combined = dateLine + "\n" + customText.Trim();

        RichTextOptions options = new(font)
        {
            WrappingLength = size - margin * 2,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top,
        };

        var attempts = 0;
        const int maxAttempts = 50;
        var measure = TextMeasurer.Measure(combined, options);

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
            measure = TextMeasurer.Measure(combined, options);
            attempts++;
        }

        using var image = new Image<Rgba32>(size, size);
        image.Mutate(ctx => ctx.Fill(backgroundColor ?? Color.White));

        var startY = (size - measure.Height) / 2f; // center vertically
        options.Origin = new PointF(size / 2f, startY);

        image.Mutate(ctx => ctx.DrawText(options, combined, textColor ?? Color.Black));

        using var ms = new MemoryStream();
        image.Save(ms, new PngEncoder());
        return ms.ToArray();
    }
}
