using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocxTemplater.Images;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;

namespace DocxTemplater.Test
{
    internal class ImageFormatterTest
    {
        [TestCase("jpg", "testImage")]
        [TestCase("tiff", "testImage")]
        [TestCase("png", "testImage")]
        [TestCase("png", "testImage_rot")]
        [TestCase("bmp", "testImage")]
        [TestCase("gif", "testImage")]
        public void ProcessTemplateWithDifferentImageTypes(string extension, string image)
        {
            var imageBytes = File.ReadAllBytes($"Resources/{image}.jpg");
            var img = Image.Load(imageBytes);
            Assert.That(img.Configuration.ImageFormatsManager.TryFindFormatByFileExtension(extension, out var format));
            var memStream = new MemoryStream();
            img.Save(memStream, format);
            imageBytes = memStream.ToArray();

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
            var img = Image.Load(imageBytes);
            img.Mutate(x => x.Resize(img.Width * 10, img.Height * 10));

            var bigImgStream = new MemoryStream();
            img.SaveAsJpeg(bigImgStream);
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
