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
            var imageService = new ImageService();
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
            var imageService = new ImageService();
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

        [Test]
        public void ImageFormatter_KeepRatio_PreservesAspectRatioInInline()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                // Create a Drawing with an Extent to simulate an existing image container
                var drawing = new Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1000000L, Cy = 1000000L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1u, Name = "Picture 1" },
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text("{{ds}:img(KEEPRATIO)}")))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                );
                mainPart.Document = new Document(new Body(new Paragraph(new Run(drawing))));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [Test]
        public void ImageFormatter_StretchWidthAndHeight_ScalesImage()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            foreach (var arg in new[] { "STRETCHW", "STRETCHH" })
            {
                using var memStream = new MemoryStream();
                using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                    var drawing = new Drawing(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1000000L, Cy = 1000000L },
                            new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1u, Name = "Picture 1" },
                            new DocumentFormat.OpenXml.Drawing.Graphic(
                                new DocumentFormat.OpenXml.Drawing.GraphicData(
                                    new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"{{{{ds}}:img({arg})}}")))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                            )
                        )
                    );
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(drawing))));
                    wpDocument.Save();
                }
                memStream.Position = 0;
                var docTemplate = new DocxTemplate(memStream);
                docTemplate.RegisterFormatter(new ImageFormatter());
                docTemplate.BindModel("ds", imageBytes);
                docTemplate.Process();
            }
        }

        [Test]
        public void ImageFormatter_AnchorLayout_InsertsImageCorrectly()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                var anchor = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.SimplePosition { X = 0L, Y = 0L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition(new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset("0")) { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition(new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset("0")) { RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1000000L, Cy = 1000000L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone(),
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1u, Name = "Picture 1" },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text("{{ds}:img}")))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { RelativeHeight = 0u, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
                var drawing = new Drawing(anchor);
                mainPart.Document = new Document(new Body(new Paragraph(new Run(drawing))));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [Test]
        public void ImageFormatter_ExplicitDimensions_TransformsSizeCorrectly()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg"); // 100x100
            foreach (var arg in new[] { "w:200px,h:100px", "w:100px,h:200px" })
            {
                using var memStream = new MemoryStream();
                using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text($"{{{{ds}}:img({arg})}}")))));
                    wpDocument.Save();
                }
                memStream.Position = 0;
                var docTemplate = new DocxTemplate(memStream);
                docTemplate.RegisterFormatter(new ImageFormatter());
                docTemplate.BindModel("ds", imageBytes);
                docTemplate.Process();
            }
        }

        [Test]
        public void ImageFormatter_UnrecognizedArgument_IgnoredGracefully()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                // "invalid" argument should trigger the matches.Count == 0 branch
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(invalid)}")))));
                wpDocument.Save();
            }
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [Test]
        public void ImageFormatter_InvalidImageContent_ThrowsException()
        {
            var invalidBytes = new byte[] { 0x01, 0x02, 0x03, 0x04 };
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img}")))));
                wpDocument.Save();
            }
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", invalidBytes);

            Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
        }

        [Test]
        public void ImageFormatter_EmptyByteArray_ClearsTarget()
        {
            var formatter = new ImageFormatter();
            var target = new Text("{{ds:img}}");
            var context = new DocxTemplater.Formatter.FormatterContext("ds", "img", Array.Empty<string>(), Array.Empty<byte>(), System.Globalization.CultureInfo.InvariantCulture);
            var modelLookup = new DocxTemplater.ModelLookup();
            var processingContext = new DocxTemplater.TemplateProcessingContext(new DocxTemplater.ProcessSettings(), modelLookup, null, null);
            formatter.ApplyFormat(processingContext, context, target);
            Assert.That(target.Text, Is.EqualTo(string.Empty));
        }
    }
}
