using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocxTemplater.Images;
using SkiaSharp;

namespace DocxTemplater.Test
{
    internal class ImageFormatterTest
    {
        private static SKEncodedImageFormat GetSkiaFormatByExtension(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                "jpg" or "jpeg" => SKEncodedImageFormat.Jpeg,
                "png" => SKEncodedImageFormat.Png,
                _ => throw new NotSupportedException($"Unsupported image extension: {extension}")
            };
        }

        [TestCase("jpg", "testImage")]
        [TestCase("png", "testImage")]
        [TestCase("png", "testImage_rot")]
        public void ProcessTemplateWithDifferentImageTypes(string extension, string image)
        {
            var imageBytes = File.ReadAllBytes($"Resources/{image}.jpg");
            using var bitmap = SKBitmap.Decode(imageBytes);
            Assert.That(bitmap, Is.Not.Null);
            var skFormat = GetSkiaFormatByExtension(extension);
            using var encoded = bitmap.Encode(skFormat, 90);
            imageBytes = encoded.ToArray();

            using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", new { MyLogo = imageBytes, EmptyArray = Array.Empty<byte>(), NullValue = (byte[])null });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        [Test]
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
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }


        [TestCase("w:14cm,h:3cm")]
        [TestCase("w:14cm")]
        [TestCase("h:1cm, r:90")]
        [TestCase("w:1cm")]
        [TestCase("h:1cm")]
        [TestCase("h:15mm")]
        public void InsertHugeImageInsertWithoutContainerFitsToPage(string argument)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            // change the size to be bigger than the page
            using var originalBitmap = SKBitmap.Decode(imageBytes);
            var newWidth = originalBitmap.Width * 10;
            var newHeight = originalBitmap.Height * 10;
            using var resized = originalBitmap.Resize(new SKImageInfo(newWidth, newHeight), new SKSamplingOptions(SKFilterMode.Linear));
            using var bigImgStream = new MemoryStream();
            using (var encodedData = resized.Encode(SKEncodedImageFormat.Jpeg, 90))
            {
                encodedData.SaveTo(bigImgStream);
            }
            imageBytes = bigImgStream.ToArray();

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(" + argument + ")}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }
    }
}
