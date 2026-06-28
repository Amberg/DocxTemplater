using DocxTemplater.ImageBase;

namespace DocxTemplater.Images
{
    /// <summary>
    /// Library-neutral image metadata required by DocxTemplater to size and embed an image.
    /// </summary>
    public sealed record ImageMetadata(int PixelWidth, int PixelHeight, ImageFormat Format, ImageRotation ExifRotation);
}
