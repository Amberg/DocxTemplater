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
                new Paragraph(new Run(new Text("{{#case: 'A'}} Match A {{/case}}"))),
                new Paragraph(new Run(new Text("{{#c: 'B'}} Match B {{/c}}"))),
                new Paragraph(new Run(new Text("{{#default}} Match Default {{/d}}"))),
                new Paragraph(new Run(new Text("{{/switch}}")))
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
                new Paragraph(new Run(new Text("{{#c: 1}} Match 1 {{/c}}"))),
                new Paragraph(new Run(new Text("{{#c: 2}} Match 2 {{/c}}"))),
                new Paragraph(new Run(new Text("{{#d}} Match Default {{/d}}"))),
                new Paragraph(new Run(new Text("{{/s}}")))
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
                new Paragraph(new Run(new Text("{{#case: 'Monday'}} Is Monday {{/c}}"))),
                new Paragraph(new Run(new Text("{{#case: 'Tuesday'}} Is Tuesday {{/c}}"))),
                new Paragraph(new Run(new Text("{{#default}} Other Day {{/d}}"))),
                new Paragraph(new Run(new Text("{{/switch}}")))
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
    }
}
