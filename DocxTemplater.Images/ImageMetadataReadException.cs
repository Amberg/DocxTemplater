using System;

namespace DocxTemplater.Images
{
    /// <summary>
    /// Represents a failure to identify image dimensions, format or orientation in an adapter package.
    /// </summary>
    public sealed class ImageMetadataReadException : Exception
    {
        public ImageMetadataReadException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
