using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class RangeLoopTest
    {
        [Test]
        public void RangeLoopWithIntModel()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{@i:ds.Count}}"))),
                new Paragraph(new Run(new Text("Item {{i}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Count = 3 });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Item 0Item 1Item 2"));
        }

        [Test]
        public void RangeLoopWithEnumerableModel()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{@i:ds.Items}}"))),
                new Paragraph(new Run(new Text("Item {{i}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Items = new[] { "A", "B", "C", "D" } });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Item 0Item 1Item 2Item 3"));
        }

        [Test]
        public void RangeLoopWithoutIndexVariable()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{@ds.Count}}"))),
                new Paragraph(new Run(new Text("Item {{Index}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Count = 2 });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Item 0Item 1"));
        }

        [Test]
        public void RangeLoopStringParsingModel()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{@i:ds.Count}}"))),
                new Paragraph(new Run(new Text("Item {{i}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Count = "2" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Trim(), Is.EqualTo("Item 0Item 1"));
        }

        [Test]
        public void RangeLoopNonParseableString_ThrowsException()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{@i:ds.Count}}"))),
                new Paragraph(new Run(new Text("Item {{i}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.ThrowException;
            docTemplate.BindModel("ds", new { Count = "abc" });
            Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
        }

        [Test]
        public void RangeLoopNegativeCount_ProducesNoIterations()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("KeepMe"))),
                new Paragraph(new Run(new Text("{{@i:ds.Count}}"))),
                new Paragraph(new Run(new Text("Item {{i}}"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Count = -5 });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var docBody = document.MainDocumentPart.Document.Body;
            Assert.That(docBody, Is.Not.Null, "Body should not be null");
            Assert.That(docBody.InnerText.Trim(), Is.EqualTo("KeepMe"));
        }
    }
}
