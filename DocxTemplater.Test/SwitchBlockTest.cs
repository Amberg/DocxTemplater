using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class SwitchBlockTest
    {
        [Test]
        public void SwitchWithStrings()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#switch: ds.Val}}"))),
                new Paragraph(new Run(new Text("{{#case: 'A'}} Match A {{/}}"))),
                new Paragraph(new Run(new Text("{{#c: 'B'}} Match B {{/}}"))),
                new Paragraph(new Run(new Text("{{#default}} Match Default {{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = "B" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Match B"));
        }

        [Test]
        public void SwitchWithNumbers()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                new Paragraph(new Run(new Text("{{#c: 1}} Match 1 {{/}}"))),
                new Paragraph(new Run(new Text("{{#c: 2}} Match 2 {{/}}"))),
                new Paragraph(new Run(new Text("{{#d}} Match Default {{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 3 });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Match Default"));
        }

        [Test]
        public void SwitchWithEnums()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#switch: ds.Day.ToString()}}"))),
                new Paragraph(new Run(new Text("{{#case: 'Monday'}} Is Monday {{/}}"))),
                new Paragraph(new Run(new Text("{{#case: 'Tuesday'}} Is Tuesday {{/}}"))),
                new Paragraph(new Run(new Text("{{#default}} Other Day {{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Day = DayOfWeek.Monday });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Is Monday"));
        }

        [Test]
        public void SwitchWithOptionalClosingTags()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                new Paragraph(new Run(new Text("{{#c: 1}}"))),
                new Paragraph(new Run(new Text("Match 1"))),
                new Paragraph(new Run(new Text("{{#c: 2}}"))),
                new Paragraph(new Run(new Text("Match 2"))),
                new Paragraph(new Run(new Text("{{#d}}"))),
                new Paragraph(new Run(new Text("Match Default"))),
                new Paragraph(new Run(new Text("{{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 2 });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Match 2"));
        }

        [Test]
        public void SwitchWithVariableNamedCOrSInside()
        {
            // Test that {{c}} or {{s}} as variable is not confused with closing tags
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                new Paragraph(new Run(new Text("{{#c: 1}} Val is {{ds.c}} {{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 1, c = "C-Value" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Val is C-Value"));
        }

        [Test]
        public void NestedSwitchWithGenericClosingTags()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#s: ds.Outer}}"))),
                new Paragraph(new Run(new Text("{{#c: 'O1'}}"))),
                    new Paragraph(new Run(new Text("Outer1-{{#s: ds.Inner}}"))),
                    new Paragraph(new Run(new Text("{{#c: 'I1'}}Inner1{{/}}"))),
                    new Paragraph(new Run(new Text("{{/}}"))),
                new Paragraph(new Run(new Text("{{/}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Outer = "O1", Inner = "I1" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Outer1-Inner1"));
        }

        [Test]
        public void SwitchWithVariableCAsCollectionStart()
        {
            // Test if {{#c}} (without colon) is treated as collection start, not as case
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#c}}Item {{.}}{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("c", new[] { "1", "2" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Item 1Item 2"));
        }

        [TestCase("{{/switch}}")]
        [TestCase("{{/s}}")]
        [TestCase("{{/case}}")]
        [TestCase("{{/c}}")]
        [TestCase("{{/default}}")]
        [TestCase("{{/d}}")]
        public void SpecificClosingTagsThrowException(string closingTag)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                new Paragraph(new Run(new Text(closingTag)))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 1 });
            var ex = Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
            Assert.That(ex.Message, Does.Contain($"Invalid syntax '{closingTag}'. Use '{{{{/}}}}' instead."));
        }
    }
}
