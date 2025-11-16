using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace WatermarkApi.Utils;

public static class PowerPointWatermarkHelper
{
    /// <summary>
    /// Adds an image watermark to all slides in a PowerPoint presentation from a file path.
    /// </summary>
    public static byte[] AddWatermarkToPresentation(string presentationPath, string imagePath)
    {
        if (string.IsNullOrWhiteSpace(presentationPath))
            throw new ArgumentException("Presentation path cannot be null or empty.", nameof(presentationPath));
        if (!System.IO.File.Exists(presentationPath))
            throw new FileNotFoundException($"Presentation file not found: {presentationPath}");
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var fileStream = new FileStream(presentationPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        return AddWatermarkToPresentation(fileStream, imagePath);
    }

    /// <summary>
    /// Adds an image watermark to all slides in a PowerPoint presentation from a stream.
    /// </summary>
    public static byte[] AddWatermarkToPresentation(Stream presentationStream, string imagePath)
    {
        if (presentationStream == null)
            throw new ArgumentNullException(nameof(presentationStream));
        if (string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Image path cannot be null or empty.", nameof(imagePath));
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var memoryStream = new MemoryStream();
        presentationStream.CopyTo(memoryStream);
        memoryStream.Position = 0;

        try
        {
            using (PresentationDocument pres = PresentationDocument.Open(memoryStream, true))
            {
                var presentationPart = pres.PresentationPart 
                    ?? throw new InvalidOperationException("Presentation part not found.");

                // Procesar cada diapositiva
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    var imagePart = slidePart.AddImagePart("image/png");
                    using (var imageStream = new FileStream(imagePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        imagePart.FeedData(imageStream);
                    }

                    AddWatermarkToSlide(slidePart, imagePart);
                }
            }

            return memoryStream.ToArray();
        }
        finally
        {
            memoryStream?.Dispose();
        }
    }

    private static void AddWatermarkToSlide(SlidePart slidePart, ImagePart imagePart)
    {
        var slide = slidePart.Slide;
        var shapeTree = slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            return;

        // Dimensiones en EMU (English Metric Units)
        // 9 pulgadas x 6.75 pulgadas
        long imageWidth = 1229600;
        long imageHeight = 1172200;

        var relationshipId = slidePart.GetIdOfPart(imagePart);
        uint newId = (uint)shapeTree.Count() + 1;

        // Crear elemento picture
        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = newId, Name = $"Watermark {newId}" },
                new P.NonVisualPictureDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = imageWidth, Cy = imageHeight }),
                new A.PresetGeometry(new A.AdjustValueList())
                { Preset = A.ShapeTypeValues.Rectangle }));

        shapeTree.AppendChild(picture);
    }
}
