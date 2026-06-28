namespace DocxTemplater.Images.Bcl
{
    /// <summary>
    /// Image formatter configured with the dependency-free .NET BCL metadata reader.
    /// </summary>
    public sealed class BclImageFormatter : ImageFormatter
    {
        public BclImageFormatter()
            : base(new BclImageMetadataReader())
        {
        }
    }
}
