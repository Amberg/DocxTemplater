namespace DocxTemplater.Images
{
    /// <summary>
    /// Reads image metadata without coupling the formatter to a specific image processing library.
    /// Adapter packages implement this interface for ImageSharp, SkiaSharp, Magick.NET or another library.
    /// </summary>
    public interface IImageMetadataReader
    {
        ImageMetadata Read(byte[] imageBytes);
    }
}
