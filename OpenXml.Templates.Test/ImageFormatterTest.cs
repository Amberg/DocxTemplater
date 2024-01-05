using OpenXml.Templates.Images;

namespace OpenXml.Templates.Test
{
    internal class ImageFormatterTest
    {

        // TODO: Test different image types
        [Test]
        public void ProcessTemplateWithDifferentImageTypes()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");
            using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.AddModel("ds", new {MyLogo = imageBytes});

            var result = docTemplate.Process();
            result.SaveAsFileAndOpenInWord();

        }
    }
}
