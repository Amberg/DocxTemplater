using DocumentFormat.OpenXml;

namespace DocxTemplater.ImageBase
{
    public sealed class ImageInformation
    {
        public ImageInformation(int pixelWidth, int pixelHeight, string imagePartRelationId, ImageRotation exifRotation, bool isSvg = false)
        {
            PixelWidth = pixelWidth;
            PixelHeight = pixelHeight;
            ImagePartRelationId = imagePartRelationId;
            ExifRotation = exifRotation;
            IsSvg = isSvg;
        }

        public int PixelWidth { get; }
        public int PixelHeight { get; }
        public string ImagePartRelationId { get; }
        public ImageRotation ExifRotation { get; }
        public bool IsSvg { get; }
    }
    public interface IImageService
    {
        uint GetImage(OpenXmlElement root, byte[] imageBytes, out ImageInformation imageInfoInformation);

        DocumentFormat.OpenXml.Drawing.Pictures.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, ImageRotation rotation);
    }
}
