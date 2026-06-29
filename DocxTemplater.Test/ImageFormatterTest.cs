using DocxTemplater.ImageBase;
using DocxTemplater.Images;
using DocxTemplater.Test.Contracts;

namespace DocxTemplater.Test
{
    [NUnit.Framework.TestFixture]
    internal sealed class ImageFormatterServiceLifecycleTests
    {
        private sealed class DummyImageService : IImageService
        {
            public uint GetImage(DocumentFormat.OpenXml.OpenXmlElement root, byte[] imageBytes, out ImageInformation imageInfoInformation)
            {
                imageInfoInformation = null;
                throw new NotSupportedException();
            }

            public DocumentFormat.OpenXml.Drawing.Pictures.Picture CreatePicture(string impagepartRelationShipId, uint propertyId, long cx, long cy, ImageRotation rotation)
            {
                throw new NotSupportedException();
            }
        }

        [NUnit.Framework.Test]
        public void CreateImageService_ReturnsNewInstance_PerCall()
        {
            var formatter = new ImageFormatter(new CoreTestImageMetadataReader());

            var first = formatter.CreateImageService();
            var second = formatter.CreateImageService();

            NUnit.Framework.Assert.That(second, NUnit.Framework.Is.Not.SameAs(first));
        }

        [NUnit.Framework.Test]
        public void CreateImageService_UsesFactoryAndReturnsNewInstance_PerCall()
        {
            var formatter = new ImageFormatter(() => new DummyImageService());

            var first = formatter.CreateImageService();
            var second = formatter.CreateImageService();

            NUnit.Framework.Assert.That(second, NUnit.Framework.Is.Not.SameAs(first));
        }

        [NUnit.Framework.Test]
        public void Constructor_WithCustomIImageServiceInstance_ThrowsImageServiceFactoryRequiredException()
        {
            NUnit.Framework.Assert.That(
                () => new ImageFormatter(new DummyImageService()),
                NUnit.Framework.Throws.TypeOf<ImageServiceFactoryRequiredException>());
        }
    }

    internal sealed class CoreImageFormatterContractTests : ImageFormatterContractTests
    {
        protected override DocxTemplater.Formatter.IFormatter CreateFormatter()
        {
            return new ImageFormatter(new CoreTestImageMetadataReader());
        }
    }

    internal sealed class CoreTestImageMetadataReader : IImageMetadataReader
    {
        public ImageMetadata Read(byte[] imageBytes)
        {
            if (imageBytes is null || imageBytes.Length < 4)
            {
                throw new ImageMetadataReadException("Image bytes are invalid.", new ArgumentException("Image bytes are invalid."));
            }

            if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8)
            {
                return new ImageMetadata(100, 100, ImageFormat.Jpeg, ImageRotation.CreateFromUnits(0));
            }

            throw new ImageMetadataReadException("Unsupported image format for core tests.", new ArgumentException("Unsupported image format for core tests."));
        }
    }
}
