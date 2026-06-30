using System.Globalization;
using DocxTemplater.Extensions;
using DocxTemplater.Formatter;
using DocxTemplater.ImageBase;
using DocxTemplater.Images;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test.Contracts
{
    public sealed record ImageRasterCase(string Extension, string ImageName);

    internal abstract class ImageFormatterContractTests
    {
        protected abstract IFormatter CreateFormatter();

        protected virtual byte[] GetDefaultRasterImageBytes()
        {
            return File.ReadAllBytes("Resources/testImage.jpg");
        }

        protected virtual byte[] GetLargeRasterImageBytes()
        {
            return GetDefaultRasterImageBytes();
        }

        [NUnit.Framework.Test]
        public void InsertSVGAndScaleAndRotate()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.svg");
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(h:1cm, r:90)}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
            docTemplate.Validate();
        }

        [NUnit.Framework.TestCase("w:14cm,h:3cm")]
        [NUnit.Framework.TestCase("w:14cm")]
        [NUnit.Framework.TestCase("h:1cm, r:90")]
        [NUnit.Framework.TestCase("w:1cm")]
        [NUnit.Framework.TestCase("h:1cm")]
        [NUnit.Framework.TestCase("h:15mm")]
        public void InsertHugeImageInsertWithoutContainerFitsToPage(string argument)
        {
            var imageBytes = GetLargeRasterImageBytes();

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(" + argument + ")}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
            docTemplate.Validate();
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_KeepRatio_PreservesAspectRatioInInline()
        {
            var imageBytes = GetDefaultRasterImageBytes();
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
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_StretchWidthAndHeight_ScalesImage()
        {
            var imageBytes = GetDefaultRasterImageBytes();
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
                docTemplate.RegisterFormatter(CreateFormatter());
                docTemplate.BindModel("ds", imageBytes);
                docTemplate.Process();
            }
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_AnchorLayout_InsertsImageCorrectly()
        {
            var imageBytes = GetDefaultRasterImageBytes();
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
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_ExplicitDimensions_TransformsSizeCorrectly()
        {
            var imageBytes = GetDefaultRasterImageBytes();
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
                docTemplate.RegisterFormatter(CreateFormatter());
                docTemplate.BindModel("ds", imageBytes);
                docTemplate.Process();
            }
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_UnrecognizedArgument_IgnoredGracefully()
        {
            var imageBytes = GetDefaultRasterImageBytes();
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(invalid)}")))));
                wpDocument.Save();
            }

            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", imageBytes);
            docTemplate.Process();
        }

        [NUnit.Framework.Test]
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
            docTemplate.RegisterFormatter(CreateFormatter());
            docTemplate.BindModel("ds", invalidBytes);

            NUnit.Framework.Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
        }

        [NUnit.Framework.Test]
        public void ImageFormatter_EmptyByteArray_ClearsTarget()
        {
            var formatter = CreateFormatter();
            var target = new Text("{{ds:img}}");
            var context = new FormatterContext("ds", "img", Array.Empty<string>(), Array.Empty<byte>(), CultureInfo.InvariantCulture);
            var processingContext = new NoopTemplateProcessingContext();
            formatter.ApplyFormat(processingContext, context, target);
            NUnit.Framework.Assert.That(target.Text, NUnit.Framework.Is.EqualTo(string.Empty));
        }

        private sealed class NoopTemplateProcessingContext : ITemplateProcessingContext
        {
            public ProcessSettings ProcessSettings => new();
            public MainDocumentPart MainDocumentPart => null;
            public IScriptCompiler ScriptCompiler => null;
            public IModelLookup ModelLookup => null;
            public IVariableReplacer VariableReplacer => null;
            public IReadOnlyCollection<ITemplateProcessorExtension> Extensions => Array.Empty<ITemplateProcessorExtension>();
            public IImageService ImageService => null;
        }
    }

    internal abstract class ImageFormatterAdapterContractTests : ImageFormatterContractTests
    {
        protected abstract IEnumerable<ImageRasterCase> RasterCases { get; }

        protected abstract byte[] ConvertToFormat(byte[] sourceImageBytes, string extension);

        [NUnit.Framework.Test]
        public void ProcessTemplateWithDifferentImageTypes()
        {
            foreach (var testCase in RasterCases)
            {
                var sourceBytes = File.ReadAllBytes($"Resources/{testCase.ImageName}.jpg");
                var imageBytes = ConvertToFormat(sourceBytes, testCase.Extension);

                using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
                var docTemplate = new DocxTemplate(fileStream);
                docTemplate.RegisterFormatter(CreateFormatter());
                docTemplate.BindModel("ds", new { MyLogo = imageBytes, EmptyArray = Array.Empty<byte>(), NullValue = (byte[])null });

                docTemplate.Process();
                docTemplate.Validate();
            }
        }
    }

    internal abstract class ImageMetadataReaderAdapterContractTests
    {
        protected abstract IImageMetadataReader CreateReader();

        protected abstract IEnumerable<ImageRasterCase> RasterCases { get; }

        protected abstract byte[] ConvertToFormat(byte[] sourceImageBytes, string extension);

        [NUnit.Framework.Test]
        public void Read_ValidImage_ReturnsExpectedMetadata()
        {
            var reader = CreateReader();
            foreach (var testCase in RasterCases)
            {
                var sourceBytes = File.ReadAllBytes($"Resources/{testCase.ImageName}.jpg");
                var imageBytes = ConvertToFormat(sourceBytes, testCase.Extension);
                var metadata = reader.Read(imageBytes);

                NUnit.Framework.Assert.That(metadata.PixelWidth, NUnit.Framework.Is.GreaterThan(0));
                NUnit.Framework.Assert.That(metadata.PixelHeight, NUnit.Framework.Is.GreaterThan(0));
                NUnit.Framework.Assert.That(metadata.Format, NUnit.Framework.Is.EqualTo(MapExpectedImageFormat(testCase.Extension)));
            }
        }

        [NUnit.Framework.Test]
        public void Read_InvalidImage_ThrowsImageMetadataReadException()
        {
            var reader = CreateReader();
            NUnit.Framework.Assert.Throws<ImageMetadataReadException>(() => reader.Read([0x00, 0x01, 0x02]));
        }

        [NUnit.Framework.Test]
        public void Read_RotatedImage_PreservesExifRotation()
        {
            var reader = CreateReader();
            var sourceBytes = File.ReadAllBytes("Resources/testImage_rot.jpg");
            var metadata = reader.Read(sourceBytes);

            NUnit.Framework.Assert.That(metadata.ExifRotation.Units, NUnit.Framework.Is.Not.EqualTo(0));
        }

        private static ImageFormat MapExpectedImageFormat(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                "jpg" or "jpeg" => ImageFormat.Jpeg,
                "png" => ImageFormat.Png,
                "gif" => ImageFormat.Gif,
                "bmp" => ImageFormat.Bmp,
                "tiff" or "tif" => ImageFormat.Tiff,
                _ => throw new ArgumentOutOfRangeException(nameof(extension), extension, "Unknown image extension")
            };
        }
    }
}