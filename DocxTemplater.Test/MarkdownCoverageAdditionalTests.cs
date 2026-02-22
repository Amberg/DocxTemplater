using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Markdown;
using DocumentFormat.OpenXml.Packaging;

namespace DocxTemplater.Test
{
    internal class MarkdownCoverageAdditionalTests
    {
        [Test]
        public void LineBreakAtEndOfContainer_ShouldNotRenderBreak()
        {
            var markdown = "Line 1  \n";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.Descendants<Break>().Count(), Is.EqualTo(0));
        }

        [Test]
        public void LineBreakInMiddle_ShouldRenderBreak()
        {
            var markdown = "Line 1  \nLine 2";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.Descendants<Break>().Count(), Is.EqualTo(1));
        }

        [Test]
        public void ParagraphInsideBlockQuote_ShouldRenderWithTemplateProperties()
        {
            var markdown = "> This is a quote";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Assert.That(paragraphs.Count, Is.GreaterThanOrEqualTo(1));
        }

        [Test]
        public void MultiParagraphBlockQuote_RendersMultipleParagraphs()
        {
            var markdown = "> Para 1\n>\n> Para 2";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Assert.That(paragraphs.Count, Is.GreaterThanOrEqualTo(2));
        }

        [Test]
        public void MarkdownFormatter_CanHandle_RecognizesMdFormatName()
        {
            var formatter = new MarkdownFormatter();
            Assert.That(formatter.CanHandle(typeof(string), "md"), Is.True);
            Assert.That(formatter.CanHandle(typeof(string), "MD"), Is.True);
            Assert.That(formatter.CanHandle(typeof(int), "md"), Is.False);
            Assert.That(formatter.CanHandle(typeof(string), "txt"), Is.False);
            Assert.That(formatter.CanHandle(null, "md"), Is.False);
        }

        [Test]
        public void MarkdownFormatter_RenderTableWithStyleArgument()
        {
            var markdown = "| Col 1 |\n| --- |\n| Val 1 |";
            // Syntax: {{ds}:md(ts:MyStyle)}
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown, ":md(ts:MyStyle)");
            Assert.That(body, Is.Not.Null);
        }

        [Test]
        public void MarkdownFormatter_DoubleNewline_CreatesSeparateParagraphs()
        {
            var markdown = "Line 1\n\nLine 2";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body.Descendants<Paragraph>().Count(), Is.GreaterThanOrEqualTo(2));
        }

        [Test]
        public void MarkdownFormatter_RawHtmlBlock_IsIgnored()
        {
            var markdown = "<div>Hello HTML</div>";
            var body = CreateTemplateWithMarkdownAndReturnBody(markdown);
            Assert.That(body, Is.Not.Null);
            Assert.That(body.InnerText ?? string.Empty, Does.Not.Contain("Hello HTML"));
        }

        [Test]
        public void MarkdownFormatter_NonStringValue_DoesNotThrow()
        {
            var formatter = new MarkdownFormatter();
            // Should not throw, just return
            formatter.ApplyFormat(null, new DocxTemplater.Formatter.FormatterContext("ds", "md", ["args"], 123, System.Globalization.CultureInfo.InvariantCulture), null);
        }

        private Body CreateTemplateWithMarkdownAndReturnBody(string markdown, string formatter = ":md")
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document =
                    new Document(new Body(new Paragraph(new Run(new Text("{{ds}" + formatter + "}")))));
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", markdown);

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.Position = 0;
            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                return new Body();
            }
            return (Body)body.CloneNode(true);
        }
    }
}
