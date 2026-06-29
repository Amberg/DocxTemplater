using System;
using DocxTemplater.ImageBase;

namespace DocxTemplater.Images
{
    /// <summary>
    /// Thrown when a custom <see cref="IImageService"/> instance is passed directly instead of a per-run factory.
    /// </summary>
    public sealed class ImageServiceFactoryRequiredException : Exception
    {
        public ImageServiceFactoryRequiredException(string message)
            : base(message)
        {
        }
    }
}