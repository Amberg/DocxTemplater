using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// With <see cref="BindingErrorHandling.HighlightErrorsInDocument"/> a template with syntax errors is not rendered:
    /// the offending tags are highlighted and the errors are listed at the top of the document instead of throwing.
    /// </summary>
    internal class HighlightSyntaxErrorsTest
    {
        private static DocxTemplate BuildTemplate(ProcessSettings settings, string[] body, string header = null, string footer = null)
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = wp.AddMainDocumentPart();
                main.Document = new Document(new Body(body.Select(p => new Paragraph(new Run(new Text(p))))));
                if (header != null)
                {
                    main.AddNewPart<HeaderPart>().Header = new Header(new Paragraph(new Run(new Text(header))));
                }
                if (footer != null)
                {
                    main.AddNewPart<FooterPart>().Footer = new Footer(new Paragraph(new Run(new Text(footer))));
                }
                wp.Save();
            }
            stream.Position = 0;
            var template = new DocxTemplate(stream, settings);
            template.BindModel("ds", new { Name = "Bob", Items = new[] { "a", "b" } });
            return template;
        }

        private static ProcessSettings Highlight()
        {
            return new ProcessSettings { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument, UiCulture = new CultureInfo("en-US") };
        }

        private static IEnumerable<Text> HighlightedTexts(OpenXmlCompositeElement root)
        {
            return root.Descendants<Run>()
                .Where(r => r.RunProperties?.GetFirstChild<Shading>()?.Fill?.Value == "FF0000")
                .SelectMany(r => r.Descendants<Text>());
        }

        [Test]
        public void SyntaxErrors_AreHighlightedAndListed_InsteadOfThrowing()
        {
            using var template = BuildTemplate(Highlight(), ["Hello {{ds.Name}}", "{{#ds.Items}}{{.}}", "{{/}} {{/}}"]);

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                // the error list comes first, then the untouched template
                Assert.That(body.InnerText, Does.StartWith("Template syntax errors:"));
                Assert.That(body.InnerText, Does.Contain("Body: '{{/}}' has no matching opening tag"));
                // nothing was rendered - the placeholders and tags are still there
                Assert.That(body.InnerText, Does.Contain("Hello {{ds.Name}}"));
                Assert.That(body.InnerText, Does.Contain("{{#ds.Items}}{{.}}"));
                Assert.That(body.InnerText, Does.Not.Contain("Bob"));
                // only the offending tag is highlighted
                Assert.That(HighlightedTexts(body).Select(x => x.Text), Is.EqualTo(["{{/}}"]));
            });
        }

        [Test]
        public void SyntaxErrorInHeader_NothingIsRendered()
        {
            using var template = BuildTemplate(Highlight(), ["Hello {{ds.Name}}"], header: "{{#ds.Items}}");

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            var header = document.MainDocumentPart.HeaderParts.Single().Header;
            Assert.Multiple(() =>
            {
                Assert.That(body.InnerText, Does.Contain("Header: '{{#ds.Items}}' is not closed"));
                Assert.That(body.InnerText, Does.Contain("Hello {{ds.Name}}"), "a valid body is not rendered either");
                Assert.That(HighlightedTexts(header).Select(x => x.Text), Is.EqualTo(["{{#ds.Items}}"]));
                Assert.That(HighlightedTexts(body), Is.Empty);
            });
        }

        [Test]
        public void TagSplitAcrossRuns_IsHighlightedAsWhole()
        {
            var settings = Highlight();
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                wp.AddMainDocumentPart().Document = new Document(new Body(new Paragraph(
                    new Run(new Text("before {{#ds.It")),
                    new Run(new Text("ems}} after")))));
                wp.Save();
            }
            stream.Position = 0;
            using var template = new DocxTemplate(stream, settings);

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(HighlightedTexts(body).Select(x => x.Text), Is.EqualTo(["{{#ds.Items}}"]));
                Assert.That(body.Descendants<Paragraph>().Last().InnerText, Is.EqualTo("before {{#ds.Items}} after"));
            });
        }

        [Test]
        public void ErrorList_IsLocalized()
        {
            var settings = Highlight();
            settings.UiCulture = new CultureInfo("fr-CH");
            settings.ErrorMessages = new TemplateErrorMessages().AddLanguage(new CultureInfo("fr"), new Dictionary<TemplateErrorCode, string>
            {
                [TemplateErrorCode.TemplateSyntaxErrors] = "Erreurs de syntaxe :\n{0}",
                [TemplateErrorCode.BlockNotClosed] = "'{0}' n'est pas fermé"
            });
            using var template = BuildTemplate(settings, ["{{#ds.Items}}"]);

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Does.StartWith("Erreurs de syntaxe :Body: '{{#ds.Items}}' n'est pas fermé"));
        }

        [Test]
        public void WarningsOnly_TemplateIsRendered()
        {
            // a malformed tag is only a warning: it stays as text and the rest renders normally
            using var template = BuildTemplate(Highlight(), ["Hello {{ds.Name}} {{broken"]);

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Is.EqualTo("Hello Bob {{broken"));
        }

        [TestCase(BindingErrorHandling.ThrowException)]
        [TestCase(BindingErrorHandling.SkipBindingAndRemoveContent)]
        public void OtherModes_StillThrow(BindingErrorHandling mode)
        {
            using var template = BuildTemplate(new ProcessSettings { BindingErrorHandling = mode }, ["{{#ds.Items}}"]);

            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
        }

        [Test]
        public void BindingErrors_StillHighlighted_WhenSyntaxIsValid()
        {
            using var template = BuildTemplate(Highlight(), ["{{ds.Missing}}"]);

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(body.InnerText, Does.Contain("Property 'Missing' not found"));
                Assert.That(HighlightedTexts(body).Select(x => x.Text), Is.EqualTo(["{{ds.Missing}}"]));
            });
        }
    }
}
