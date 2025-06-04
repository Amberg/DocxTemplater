using DocumentFormat.OpenXml;

namespace DocxTemplater.ImageBase
{
    public sealed record ImageInformation(int PixelWidth, int PixelHeight, string ImagePartRelationId, ImageRotation ExifRotation, bool IsSvg = false);
    public interface IImageService
    {
        uint GetImage(OpenXmlElement root, byte[] imageBytes, out ImageInformation imageInfoInformation);

        DocumentFormat.OpenXml.Drawing.Pictures.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, ImageRotation rotation);
    }
}
