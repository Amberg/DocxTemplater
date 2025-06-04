using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.ImageBase;
using DocxTemplater.Images;
using System.Reflection;
using System.Text;

namespace DocxTemplater.Test
{
    [TestFixture]
    public class ImageServiceTest
    {
        private byte[] CreateSampleSvg(int width = 100, int height = 100)
        {
            string svgContent = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>
<svg width=""{width}"" height=""{height}"" xmlns=""http://www.w3.org/2000/svg"">
  <rect width=""100%"" height=""100%"" fill=""#f0f0f0""/>
  <circle cx=""{width / 2}"" cy=""{height / 2}"" r=""{Math.Min(width, height) / 4}"" fill=""blue""/>
  <text x=""{width / 2}"" y=""{height / 2 + 5}"" font-family=""Arial"" font-size=""10"" text-anchor=""middle"" fill=""white"">SVG</text>
</svg>";
            return Encoding.UTF8.GetBytes(svgContent);
        }

        private byte[] CreateSvgWithViewBox()
        {
            string svgContent = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>
<svg viewBox=""0 0 200 150"" xmlns=""http://www.w3.org/2000/svg"">
  <rect width=""100%"" height=""100%"" fill=""#f0f0f0""/>
  <circle cx=""100"" cy=""75"" r=""50"" fill=""green""/>
  <text x=""100"" y=""80"" font-family=""Arial"" font-size=""10"" text-anchor=""middle"" fill=""white"">ViewBox</text>
</svg>";
            return Encoding.UTF8.GetBytes(svgContent);
        }

        [Test]
        public void IsSvgImageWithValidSvgReturnsTrue()
        {
            // Arrange
            byte[] svgBytes = CreateSampleSvg();

            // Act
            bool isSvg = InvokePrivateMethod<bool>(typeof(ImageService), "IsSvgImage", svgBytes);

            // Assert
            Assert.That(isSvg, Is.True);
        }

        [Test]
        public void IsSvgImageWithInvalidDataReturnsFalse()
        {
            // Arrange
            byte[] nonSvgBytes = Encoding.UTF8.GetBytes("This is not an SVG file");

            // Act
            bool isSvg = InvokePrivateMethod<bool>(typeof(ImageService), "IsSvgImage", nonSvgBytes);

            // Assert
            Assert.That(isSvg, Is.False);
        }

        [Test]
        public void ExtractSvgWidthWithExplicitWidthReturnsCorrectValue()
        {
            // Arrange
            int expectedWidth = 200;
            byte[] svgBytes = CreateSampleSvg(expectedWidth, 100);

            // Act
            int? width = InvokePrivateMethod<int?>(typeof(ImageService), "ExtractSvgWidth", svgBytes);

            // Assert
            Assert.That(width, Is.EqualTo(expectedWidth));
        }

        [Test]
        public void ExtractSvgHeightWithExplicitHeightReturnsCorrectValue()
        {
            // Arrange
            int expectedHeight = 150;
            byte[] svgBytes = CreateSampleSvg(100, expectedHeight);

            // Act
            int? height = InvokePrivateMethod<int?>(typeof(ImageService), "ExtractSvgHeight", svgBytes);

            // Assert
            Assert.That(height, Is.EqualTo(expectedHeight));
        }

        [Test]
        public void ExtractSvgDimensionsWithViewBoxReturnsCorrectValues()
        {
            // Arrange
            byte[] svgBytes = CreateSvgWithViewBox();

            // Act
            int? width = InvokePrivateMethod<int?>(typeof(ImageService), "ExtractSvgWidth", svgBytes);
            int? height = InvokePrivateMethod<int?>(typeof(ImageService), "ExtractSvgHeight", svgBytes);

            // Assert
            Assert.That(width, Is.EqualTo(200));
            Assert.That(height, Is.EqualTo(150));
        }

        [Test]
        public void GetImageWithSvgImageReturnsSvgImageInformation()
        {
            // Arrange
            var imageService = new ImageService();
            byte[] svgBytes = CreateSampleSvg(200, 150);

            // Create a simple document to test with
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            var mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Test")))));
            wpDocument.Save();

            var root = mainPart.Document;

            // Act
            var propertyId = imageService.GetImage(root, svgBytes, out var imageInfo);

            // Assert
            Assert.That(imageInfo, Is.Not.Null);
            Assert.That(imageInfo.PixelWidth, Is.EqualTo(200));
            Assert.That(imageInfo.PixelHeight, Is.EqualTo(150));
            Assert.That(imageInfo.IsSvg, Is.True);
            Assert.That(imageInfo.ImagePartRelationId, Is.Not.Null.Or.Empty);

            // Verify that the SVG image part was created with the correct content type
            var imagePart = mainPart.GetPartById(imageInfo.ImagePartRelationId) as ImagePart;
            Assert.That(imagePart, Is.Not.Null);
            Assert.That(imagePart.ContentType, Is.EqualTo("image/svg+xml"));
        }

        [Test]
        public void CreatePictureWithSvgImageCreatesPictureWithCorrectProperties()
        {
            // Arrange
            var imageService = new ImageService();
            const string relationshipId = "rId123";
            const uint propertyId = 1;
            const long cx = 200 * 9525; // 200 pixels in EMUs
            const long cy = 150 * 9525; // 150 pixels in EMUs
            var rotation = ImageRotation.CreateFromDegree(0);

            // Act
            var picture = imageService.CreatePicture(relationshipId, propertyId, cx, cy, rotation);

            // Assert
            Assert.That(picture, Is.Not.Null);

            // Check if the picture has the correct relationship ID
            var blip = picture.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            Assert.That(blip, Is.Not.Null);
            Assert.That(blip.Embed.Value, Is.EqualTo(relationshipId));

            // Check dimensions
            var extents = picture.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
            Assert.That(extents, Is.Not.Null);
            Assert.That(extents.Cx.Value, Is.EqualTo(cx));
            Assert.That(extents.Cy.Value, Is.EqualTo(cy));
        }

        // Helper method to invoke private static methods via reflection
        private static T InvokePrivateMethod<T>(Type type, string methodName, params object[] parameters)
        {
            var method = type.GetMethod(methodName,
                BindingFlags.NonPublic | BindingFlags.Static) ?? throw new ArgumentException($"Method {methodName} not found on type {type.Name}");
            var result = method.Invoke(null, parameters);
            return result != null ? (T)result : default;
        }
    }
}