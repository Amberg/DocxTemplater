using DocxTemplater.ImageBase;
using DocxTemplater.Images;
using DocxTemplater.Test.Contracts;

namespace DocxTemplater.Test
{
    [NUnit.Framework.TestFixture]
    internal sealed class ImageFormatterServiceLifecycleTests
    {
        [NUnit.Framework.Test]
        public void CreateImageService_ReturnsNewInstance_PerCall()
        {
            var formatter = new ImageFormatter(new CoreTestImageMetadataReader());

            var first = formatter.CreateImageService();
            var second = formatter.CreateImageService();

            NUnit.Framework.Assert.That(second, NUnit.Framework.Is.Not.SameAs(first));
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
