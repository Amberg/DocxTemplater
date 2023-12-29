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
                    new Run(
                        new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }),
                        new Text("Replace Me")),
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
                        new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }),
                        new Text("StartText")),
                    new Run(new Text("Text after start")), 
                    new Run(new Text("Text2"))
                ),
                new Paragraph(
                    new Run(new Text("Middle 1")),
                    new Run(new Text("Middle 2"))
                ),
                new Paragraph(
                    new Run(new Text("Next Paragraph")),
                    new Run(new Text("42")),
                    new Run(
                        new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(false) }),
                        new Text("EndText")),
                    new Run(new Text("Text after end"))
                )
            ));
            wpDocument.Save();

            var characterMap = new CharacterMap(wpDocument.MainDocumentPart.Document.Body);
            var start = characterMap.Text.IndexOf("StartText");
            var end = characterMap.Text.IndexOf("EndText");
            var elements = characterMap.CutBetween(characterMap[start].Element, characterMap[end].Element);
            Assert.That(elements.Count, Is.EqualTo(3));
            CollectionAssert.AreEqual(new[]{ "Text after start", "Text2", "Middle 1", "Middle 2", "Next Paragraph", "42"},elements.SelectMany(x => x.Descendants<Text>().Select(x => x.InnerText)));

            // check EndText and StartText is still in the document
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Any(x => x.Text == "StartText"), Is.True);
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Any(x => x.Text == "EndText"), Is.True);

            // check that Text after end and Some leading text is preserved
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Any(x => x.Text == "Text after end"), Is.True);
            Assert.That(wpDocument.MainDocumentPart.Document.Body.Descendants<Text>().Any(x => x.Text == "Some leading text"), Is.True);
        }
    }
}