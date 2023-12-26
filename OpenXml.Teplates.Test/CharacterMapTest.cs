using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OpenXml.Templates;

namespace OpenXml.Teplates.Test
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void ReplaceTextBoldIsPreserved()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new Run(new Text("This Value:")),
                    new Run(new Text("Replace Me"), new Bold(){Val = OnOffValue.FromBoolean(true)}),
                    new Run(new Text("will be replaced"))
                    )));
                wpDocument.Save();
                
            var characterMap = new CharacterMap(wpDocument.MainDocumentPart.Document.Body);
            characterMap.ReplaceText("Replace Me", "Replaced");

            // check that bold is preserved
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Bold>().First().Val, Is.EqualTo(OnOffValue.FromBoolean(true)));
            // check that text is replaced
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Skip(1).First().Text, Is.EqualTo("Replaced"));

        }
    }
}