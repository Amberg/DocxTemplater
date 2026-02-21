using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class CoreCoverageAdditionalTests
    {
        [Test]
        public void RangeLoop_MissingVariable_ShowsErrorInDocument()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{@i:MissingVar}}"))),
                    new Paragraph(new Run(new Text("Item {{i}}"))),
                    new Paragraph(new Run(new Text("{{/}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument;
            var result = docTemplate.Process();

            using var resultDoc = WordprocessingDocument.Open(result, false);
            var body = resultDoc.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Does.Contain("Model MissingVar not found"));
        }

        [Test]
        public void RangeLoop_NonIntegerString_ShowsErrorInDocument()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{@i:ds.Count}}"))),
                    new Paragraph(new Run(new Text("Item {{i}}"))),
                    new Paragraph(new Run(new Text("{{/}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument;
            docTemplate.BindModel("ds", new { Count = "NotAnInt" });
            var result = docTemplate.Process();

            using var resultDoc = WordprocessingDocument.Open(result, false);
            var body = resultDoc.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Does.Contain("is not an integer"));
        }

        [Test]
        public void RangeLoop_UnsupportedValueType_ThrowsException()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{@i:ds.Val}}"))),
                    new Paragraph(new Run(new Text("Item {{i}}"))),
                    new Paragraph(new Run(new Text("{{/}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.ThrowException;
            docTemplate.BindModel("ds", new { Val = System.DateTime.Now });
            Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
        }

        [Test]
        public void RangeLoop_PlainEnumerable_ExpandsCorrectly()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{@i:ds.Items}}"))),
                    new Paragraph(new Run(new Text("Item {{i}}"))),
                    new Paragraph(new Run(new Text("{{/}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            // Use a plain Enumerable that doesn't implement ICollection/IReadOnlyCollection
            var items = Enumerable.Range(0, 3).Select(x => x);
            docTemplate.BindModel("ds", new { Items = items });
            var result = docTemplate.Process();

            using var resultDoc = WordprocessingDocument.Open(result, false);
            var body = resultDoc.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("Item 0Item 1Item 2"));
        }

        [Test]
        public void Switch_WithoutCaseChildren_DoesNotThrow()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{#switch: ds.Val}}{{/switch}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 1 });
            docTemplate.Process();
            // Should not throw
        }

        [Test]
        public void Switch_UnresolvableCaseExpression_ShowsErrorInDocument()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                    new Paragraph(new Run(new Text("{{#c: ds.NonExistentProp}}"))),
                    new Paragraph(new Run(new Text("Match"))),
                    new Paragraph(new Run(new Text("{{/s}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.Settings.BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument;
            docTemplate.BindModel("ds", new { Val = 1 });
            var result = docTemplate.Process();

            using var resultDoc = WordprocessingDocument.Open(result, false);
            Assert.That(resultDoc.MainDocumentPart.Document.Body.InnerText, Does.Contain("NonExistentProp"));
        }

        [Test]
        public void Switch_NoMatchAndNoDefault_ProducesEmptyOutput()
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("{{#s: ds.Val}}"))),
                    new Paragraph(new Run(new Text("{{#c: 2}} Match 2 {{/c}}"))),
                    new Paragraph(new Run(new Text("{{/s}}")))
                ));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Val = 1 });
            var result = docTemplate.Process();

            using var resultDoc = WordprocessingDocument.Open(result, false);
            Assert.That(resultDoc.MainDocumentPart.Document.Body.InnerText.Trim(), Is.Empty);
        }
    }
}
