using System.Text;
using DocxTemplater.Images;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocxTemplater.Test
{
    internal class ImageCoverageAdditionalTests
    {
        [Test]
        public void SvgHelper_StripsUtf8BomAndLeadingWhitespace()
        {
            var svg = "   <svg width=\"100\" height=\"100\"></svg>";
            var bytes = Encoding.UTF8.GetBytes(svg);
            var bomBytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(bytes).ToArray();

            Assert.That(SvgHelper.TryReadAsSvg(bomBytes, out var w, out _), Is.True);
            Assert.That(w, Is.EqualTo(100));
        }

        [Test]
        public void SvgHelper_FallsBackToViewBoxWhenWidthHeightMissing()
        {
            var svg = "<svg viewBox=\"0 0 500 600\"></svg>";
            Assert.That(SvgHelper.TryReadAsSvg(Encoding.UTF8.GetBytes(svg), out var w, out var h), Is.True);
            Assert.That(w, Is.EqualTo(500));
            Assert.That(h, Is.EqualTo(600));
        }

        [Test]
        public void SvgHelper_ParsesPercentAndPxUnits()
        {
            var svg = "<svg width=\"100%\" height=\"200px\"></svg>";
            Assert.That(SvgHelper.TryReadAsSvg(Encoding.UTF8.GetBytes(svg), out var w, out var h), Is.True);
            Assert.That(w, Is.EqualTo(100));
            Assert.That(h, Is.EqualTo(200));
        }

        [Test]
        public void SvgHelper_InvalidXml_ReturnsFalse()
        {
            var invalid = "<svg width=\"100\"";
            Assert.That(SvgHelper.TryReadAsSvg(Encoding.UTF8.GetBytes(invalid), out _, out _), Is.False);
        }

        [Test]
        public void SvgHelper_NonSvgRootElement_ReturnsFalse()
        {
            var notSvg = "<html></html>";
            Assert.That(SvgHelper.TryReadAsSvg(Encoding.UTF8.GetBytes(notSvg), out _, out _), Is.False);
        }

        [Test]
        public void ImageService_ReusesImagePartFromCache()
        {
            var imageService = new ImageService(new DefaultImageMetadataReader());
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            var mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // First call - pass the RootElement (Document)
            imageService.GetImage(mainPart.Document, imageBytes, out var info1);
            // Second call - should hit cache
            imageService.GetImage(mainPart.Document, imageBytes, out var info2);

            Assert.That(info1.ImagePartRelationId, Is.EqualTo(info2.ImagePartRelationId));
        }

        [Test]
        public void ImageService_AddsImagePartsToHeaderAndFooter()
        {
            var imageService = new ImageService(new DefaultImageMetadataReader());
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            var mainPart = wpDocument.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header();
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer();

            imageService.GetImage(headerPart.Header, imageBytes, out _);
            imageService.GetImage(footerPart.Footer, imageBytes, out _);

            // Check SVG in header/footer too
            var svgBytes = File.ReadAllBytes("Resources/testImage.svg");
            imageService.GetImage(headerPart.Header, svgBytes, out _);
            imageService.GetImage(footerPart.Footer, svgBytes, out _);
        }

    }
}
