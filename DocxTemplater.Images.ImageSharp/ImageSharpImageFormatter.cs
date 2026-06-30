namespace DocxTemplater.Images.ImageSharp
{
    /// <summary>
    /// Image formatter configured with the default ImageSharp metadata reader.
    /// </summary>
    public sealed class ImageSharpImageFormatter : ImageFormatter
    {
        public ImageSharpImageFormatter()
            : base(new ImageSharpImageMetadataReader())
        {
        }
    }
}
