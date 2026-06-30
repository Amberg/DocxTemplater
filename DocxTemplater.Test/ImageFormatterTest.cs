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
            var formatter = new ImageFormatter(new DefaultImageMetadataReader());

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
    }

    internal sealed class CoreImageFormatterContractTests : ImageFormatterContractTests
    {
        protected override DocxTemplater.Formatter.IFormatter CreateFormatter()
        {
            return new ImageFormatter();
        }
    }
}
