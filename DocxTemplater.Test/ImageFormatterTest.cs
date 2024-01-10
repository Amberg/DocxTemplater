using DocxTemplater.Images;
using SixLabors.ImageSharp;

namespace DocxTemplater.Test
{
    internal class ImageFormatterTest
    {
        [TestCase("jpg")]
        [TestCase("tiff")]
        [TestCase("png")]
        [TestCase("bmp")]
        [TestCase("gif")]
        public void ProcessTemplateWithDifferentImageTypes(string extension)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            var img = Image.Load(imageBytes);
            Assert.That(img.Configuration.ImageFormatsManager.TryFindFormatByFileExtension(extension, out var format));
            var memStream = new MemoryStream();
            img.Save(memStream, format);
            imageBytes = memStream.ToArray();

            using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", new { MyLogo = imageBytes });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }
    }
}
