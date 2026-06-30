using System;
using DocxTemplater.ImageBase;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;

namespace DocxTemplater.Images.ImageSharp
{
    /// <summary>
    /// Reads image metadata using SixLabors.ImageSharp.
    /// </summary>
    public sealed class ImageSharpImageMetadataReader : IImageMetadataReader
    {
        public ImageMetadata Read(byte[] imageBytes)
        {
            try
            {
                using var image = Image.Load(imageBytes);
                var exifRotation = ImageRotation.CreateFromUnits(0);
                if (image.Metadata?.ExifProfile?.TryGetValue(ExifTag.Orientation, out var orientationValue) == true)
                {
                    exifRotation = ImageRotation.CreateFromExifRotation((ExifRotation)orientationValue.Value);
                }

                return new ImageMetadata(
                    image.Width,
                    image.Height,
                    MapImageFormat(image.Metadata?.DecodedImageFormat?.Name),
                    exifRotation);
            }
            catch (Exception e) when (e is InvalidImageContentException or UnknownImageFormatException)
            {
                throw new ImageMetadataReadException("Could not read image metadata using ImageSharp.", e);
            }
        }

        private static ImageFormat MapImageFormat(string imageFormatName)
        {
            return imageFormatName switch
            {
                "TIFF" => ImageFormat.Tiff,
                "BMP" => ImageFormat.Bmp,
                "GIF" => ImageFormat.Gif,
                "JPEG" => ImageFormat.Jpeg,
                "PNG" => ImageFormat.Png,
                _ => throw new ImageMetadataReadException($"Unsupported image format '{imageFormatName}'.", null)
            };
        }
    }
}
