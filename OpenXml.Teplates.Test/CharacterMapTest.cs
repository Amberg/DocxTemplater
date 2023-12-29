using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Templates;

namespace OpenXml.Templates.Test
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
                new Run(
                    new RunProperties(new Bold() {Val = OnOffValue.FromBoolean(true)}),
                    new Text("Replace Me")),
                new Run(new Text("will be replaced"))
            )));
            wpDocument.Save();

            var characterMap = new CharacterMap(wpDocument.MainDocumentPart.Document.Body);
            characterMap.ReplaceText("Replace Me", "Replaced");

            // check that bold is preserved
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Bold>().First().Val,
                Is.EqualTo(OnOffValue.FromBoolean(true)));
            // check that text is replaced
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Skip(1).First().Text,
                Is.EqualTo("Replaced"));

        }



        [Test]
        public void CutBetween()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Some leading text")),
                    new Run(
                        new RunProperties(new Bold() {Val = OnOffValue.FromBoolean(true)}),
                        new Text("StartText")),
                    new Run(new Text("Text after start"), new Text("Text2")) // two text elemnts in same run
                ),
                new Paragraph(
                    new Run(new Text("Middle 1 - Middle 1 end")),
                    new Run(new Text("Middle 2 - Middle 2 end"))
                ),
                new Paragraph(
                    new Run(new Text("Next Paragraph")),
                    new Run(new Text("42")),
                    new Run(
                        new RunProperties(new Bold() {Val = OnOffValue.FromBoolean(false)}),
                        new Text("EndText")),
                    new Run(new Text("Text after end"))
                )
            ));
            wpDocument.Save();
            memStream.Position = 0;

            Console.WriteLine(wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
        }
    }
}