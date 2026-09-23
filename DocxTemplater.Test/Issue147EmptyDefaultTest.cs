using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class Issue147EmptyDefaultTest
    {
        private static string Render(string template)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(template)))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Selector = "x", Flag = true });
            var result = docTemplate.Process();
            docTemplate.Validate();
            var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body.InnerText;
        }

        [Test]
        public void EmptyDefaultFollowedByConditionInSameParagraph()
        {
            Assert.That(Render("{{#switch: ds.Selector}}{{#case: 'x'}}X{{/}}{{#default}}{{/}}{{/}}{?{ds.Flag}}b{{/}}"), Is.EqualTo("Xb"));
        }

        [Test]
        public void EmptyDefaultFollowedByTextInSameParagraph()
        {
            Assert.That(Render("{{#switch: ds.Selector}}{{#case: 'x'}}X{{/}}{{#default}}{{/}}{{/}}tail"), Is.EqualTo("Xtail"));
        }

        [Test]
        public void NonEmptyDefaultFollowedByConditionInSameParagraph()
        {
            Assert.That(Render("{{#switch: ds.Selector}}{{#case: 'x'}}X{{/}}{{#default}}D{{/}}{{/}}{?{ds.Flag}}b{{/}}"), Is.EqualTo("Xb"));
        }

        [Test]
        public void NoDefaultFollowedByConditionInSameParagraph()
        {
            Assert.That(Render("{{#switch: ds.Selector}}{{#case: 'x'}}X{{/}}{{/}}{?{ds.Flag}}b{{/}}"), Is.EqualTo("Xb"));
        }

        [Test]
        public void EmptyCaseAndEmptyDefaultFollowedByCondition()
        {
            Assert.That(Render("{{#switch: ds.Selector}}{{#case: 'x'}}{{/}}{{#default}}{{/}}{{/}}{?{ds.Flag}}b{{/}}"), Is.EqualTo("b"));
        }

        /// <summary>
        /// Same root cause, without a switch: the block ends in the next paragraph, and that paragraph
        /// carries the End_ marker of the block while still holding the content of the following block.
        /// The anchor of the body has to be placed in front of it, not behind it.
        /// a=false is the case that shows up in the rendered output instead of only in the Debug-only
        /// block validation: with the anchor behind that paragraph the last paragraph becomes part of
        /// the body of A, which is not rendered, so "tail" is dropped.
        /// </summary>
        [TestCase(true)]
        [TestCase(false)]
        public void ConditionEndsInParagraphThatOpensTheNextCondition(bool a)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("head{?{ds.A}}"))),
                new Paragraph(new Run(new Text("{{/}}{?{ds.B}}"))),
                new Paragraph(new Run(new Text("{{/}}tail")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { A = a, B = true });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Is.EqualTo("headtail"));
        }

        [TestCase("Herr", null, "Herrn Hans Muster")]
        [TestCase("Frau", "Dr.", "Frau Dr. Hans Muster")]
        [TestCase("", "Dr.", "Dr. Hans Muster")]
        [TestCase("", null, "Hans Muster")]
        public void SalutationLineWithEmptyDefault(string salutation, string title, string expected)
        {
            const string template = "{{#switch: person.Salutation}}{{#case: 'Herr'}}Herrn {{/}}{{#case: 'Frau'}}Frau {{/}}{{#default}}{{/}}{{/}}" +
                                    "{?{person.Title != null}}{{person.Title}} {{/}}{{person.FirstName}} {{person.LastName}}";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(template)))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("person", new { Salutation = salutation, Title = title, FirstName = "Hans", LastName = "Muster" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Is.EqualTo(expected));
        }
    }
}
