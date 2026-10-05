using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class TemplateSyntaxValidationTest
    {
        private static DocxTemplate BuildTemplate(params string[] paragraphs)
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = wp.AddMainDocumentPart();
                main.Document = new Document(new Body(paragraphs.Select(p => new Paragraph(new Run(new Text(p))))));
                wp.Save();
            }
            stream.Position = 0;
            return new DocxTemplate(stream);
        }

        private static IReadOnlyList<TemplateSyntaxError> Validate(params string[] paragraphs)
        {
            using var template = BuildTemplate(paragraphs);
            return template.ValidateTemplateSyntax();
        }

        [Test]
        public void ValidTemplate_HasNoErrors()
        {
            var errors = Validate(
                "Hello {{customer.Name}:toUpper} {{(customer.Age + 1)}}",
                "{{#Items}}{{.Name}}{{:s:}}, {{/Items}}",
                "{?{customer.Age > 18 && customer.Name != 'Bob'}}adult{{:}}child{{/}}",
                "{{#switch: customer.Kind}}{{#case: 'A'}}A{{#c: 'B'}}B{{/}}{{#default}}other{{/}}{{/}}",
                "{{@i:customer.Count}}{{i}}{{/}}",
                "{{#Orders}}{{#.Lines}}{{.Name}}{{/.Lines}}{{/Orders}}",
                "{{:PageBreak}}{{:break}}{{:SectionBreak}}",
                "{{:ignore}} {{ not a tag }} {{/:ignore}}");

            Assert.That(errors, Is.Empty, string.Join(Environment.NewLine, errors));
        }

        [Test]
        public void UnclosedBlocks_AreReported()
        {
            var errors = Validate("{{#Items}}{{.Name}}", "{?{a > 1}}foo");

            Assert.That(errors.Select(x => x.Tag), Is.EqualTo(["{{#Items}}", "{?{a > 1}}"]));
            Assert.That(errors.Select(x => x.Message), Is.All.Contains("is not closed"));
        }

        [Test]
        public void ClosingTagWithoutOpening_IsReported()
        {
            var errors = Validate("{{Name}}{{/}}");

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Tag, Is.EqualTo("{{/}}"));
            Assert.That(errors[0].Message, Does.Contain("no matching opening tag"));
            Assert.That(errors[0].Part, Is.EqualTo("Body"));
            Assert.That(errors[0].Context, Does.Contain("{{Name}}"));
        }

        [TestCase("{{#Items}}foo{{/Orders}}", "{{/Orders}}", "does not match '{{#Items}}'")]
        [TestCase("{?{a}}foo{{/Items}}", "{{/Items}}", "expected '{{/}}'")]
        [TestCase("{{:ignore}}foo{{/}}", "{{/}}", "expected '{{/:ignore}}'")]
        [TestCase("{{#Items}}foo{{/:ignore}}", "{{/:ignore}}", "does not match '{{#Items}}'")]
        public void MismatchedClosingTag_IsReported(string template, string tag, string message)
        {
            var errors = Validate(template);

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Tag, Is.EqualTo(tag));
            Assert.That(errors[0].Message, Does.Contain(message));
        }

        [TestCase("{{#Items}}foo{{/}}")]
        [TestCase("{{#items}}foo{{/Items}}")]
        [TestCase("{{#ds.Items}}foo{{/Items}}")]
        [TestCase("{{#Items}}foo{{/ds.Items}}")]
        [TestCase("{{#Outer}}{{#.Inner}}foo{{/Inner}}{{/Outer}}")]
        public void LoopClosedWithEmptyOrEquivalentName_IsValid(string template)
        {
            Assert.That(Validate(template), Is.Empty);
        }

        [TestCase("foo{{:}}bar", "{{:}}", "only allowed directly inside a condition")]
        [TestCase("{{#Items}}foo{{:}}bar{{/Items}}", "{{:}}", "only allowed directly inside a condition")]
        [TestCase("{?{a}}foo{{:}}bar{{:}}baz{{/}}", "{{:}}", "more than one else")]
        [TestCase("foo{{:s:}}bar", "{{:s:}}", "only allowed directly inside a collection loop")]
        [TestCase("{{#case: 'A'}}foo{{/}}", "{{#case: 'A'}}", "must be inside a '{{#switch: ...}}' block")]
        [TestCase("{{#switch}}foo{{/}}", "{{#switch}}", "requires an expression")]
        [TestCase("{{:Foo}}", "{{:Foo}}", "Unknown keyword")]
        public void MisplacedOrInvalidTag_IsReported(string template, string tag, string message)
        {
            var errors = Validate(template);

            Assert.That(errors, Is.Not.Empty);
            Assert.That(errors[0].Tag, Is.EqualTo(tag));
            Assert.That(errors[0].Message, Does.Contain(message));
        }

        [Test]
        public void SeparatorInRangeLoop_IsReported()
        {
            var errors = Validate("{{@i:3}}{{i}}{{:s:}}{{/}}");

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Tag, Is.EqualTo("{{:s:}}"));
        }

        [TestCase("Hello {{Name}", "{{Name}", "is not terminated")]
        [TestCase("Hello {{first name}}", "{{first name}}", "Invalid tag")]
        [TestCase("Hello Name}} foo", "}}", "without matching")]
        [TestCase("{{#Items:foo}}", "{{#Items:foo}}", "Invalid tag")]
        public void MalformedTag_IsReported(string template, string tag, string message)
        {
            var errors = Validate(template);

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Tag, Is.EqualTo(tag));
            Assert.That(errors[0].Message, Does.Contain(message));
        }

        [TestCase("{?{(a > 5}}foo{{/}}", "Unclosed '('")]
        [TestCase("{?{a > 5)}}foo{{/}}", "Unbalanced ')'")]
        [TestCase("{?{a == 'foo}}foo{{/}}", "Unterminated string literal")]
        [TestCase("{?{ }}foo{{/}}", "empty expression")]
        [TestCase("{{(a + \"x)}}", "Unterminated string literal")]
        public void InvalidExpression_IsReported(string template, string message)
        {
            var errors = Validate(template);

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Message, Does.Contain(message));
        }

        [TestCase("{{/switch}}")]
        [TestCase("{{/:foo}}")]
        public void PatternMatcherError_IsReportedInsteadOfThrown(string template)
        {
            var errors = Validate(template);

            Assert.That(errors, Has.Count.EqualTo(1));
            Assert.That(errors[0].Tag, Is.EqualTo(template));
            Assert.That(errors[0].Message, Does.Contain("Invalid syntax"));
        }

        [Test]
        public void AllErrorsAreCollected_InDocumentOrder()
        {
            var errors = Validate("{{Name}", "{{:}}", "{{#Items}}", "{{/Orders}}", "{{/}}");

            Assert.That(errors.Select(x => x.Tag), Is.EqualTo(["{{Name}", "{{:}}", "{{/Orders}}", "{{/}}"]));
        }

        [Test]
        public void ErrorsInHeaderAndFooter_AreReportedWithPart()
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = wp.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("{{Name}}")))));
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Header(new Paragraph(new Run(new Text("{{#Items}}"))));
                var footer = main.AddNewPart<FooterPart>();
                footer.Footer = new Footer(new Paragraph(new Run(new Text("{{/}}"))));
                wp.Save();
            }
            stream.Position = 0;
            using var template = new DocxTemplate(stream);

            var errors = template.ValidateTemplateSyntax();

            Assert.That(errors.Select(x => (x.Part, x.Tag)), Is.EqualTo([("Header", "{{#Items}}"), ("Footer", "{{/}}")]));
        }

        [Test]
        public void TagSplitAcrossRuns_IsValid()
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = wp.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(
                    new Run(new Text("{{#Ite")),
                    new Run(new Text("ms}}{{.}}{{/Items")),
                    new Run(new Text("}}")))));
                wp.Save();
            }
            stream.Position = 0;
            using var template = new DocxTemplate(stream);

            Assert.That(template.ValidateTemplateSyntax(), Is.Empty);
        }

        [Test]
        public void Validate_DoesNotModifyDocument()
        {
            using var template = BuildTemplate("{{#Items}}{{.}}{{/Items}}");
            Assert.That(template.ValidateTemplateSyntax(), Is.Empty);

            template.BindModel("Items", new[] { "a", "b" });
            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Is.EqualTo("ab"));
        }

        [Test]
        public void Validate_AfterGetTemplateSchema_UsesOriginalText()
        {
            using var template = BuildTemplate("{{#Items}}{{.}}{{/Items}}", "{{Name}");
            template.GetTemplateSchema();

            var errors = template.ValidateTemplateSyntax();

            Assert.That(errors.Select(x => x.Tag), Is.EqualTo(["{{Name}"]));
        }

        [Test]
        public void Validate_AfterProcess_Throws()
        {
            using var template = BuildTemplate("{{Name}}");
            template.BindModel("Name", "foo");
            template.Process();

            Assert.Throws<OpenXmlTemplateException>(() => template.ValidateTemplateSyntax());
        }

        private static IEnumerable<string> ResourceTemplates()
        {
            return Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Resources"), "*.docx")
                .Select(Path.GetFileName);
        }

        [TestCaseSource(nameof(ResourceTemplates))]
        public void ExistingTestTemplates_HaveNoSyntaxErrors(string fileName)
        {
            using var template = DocxTemplate.Open(Path.Combine(AppContext.BaseDirectory, "Resources", fileName));

            var errors = template.ValidateTemplateSyntax();

            Assert.That(errors, Is.Empty, string.Join(Environment.NewLine, errors));
        }

        [Test]
        public void Process_ThrowsWithAllSyntaxErrors()
        {
            using var template = BuildTemplate("{{#Items}}{{.}}{{/Orders}}", "{{:}}", "{{/}}");
            template.BindModel("Items", new[] { "a" });

            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            Assert.That(ex.Message, Does.Contain("'{{/Orders}}' does not match '{{#Items}}'"));
            Assert.That(ex.Message, Does.Contain("'{{:}}' (else) is only allowed"));
            Assert.That(ex.Message, Does.Contain("'{{/}}' has no matching opening tag"));
        }

        [Test]
        public void Process_RendersDespiteWarnings()
        {
            using var template = BuildTemplate("{{Name} {{Name}} }}");
            var errors = template.ValidateTemplateSyntax();
            Assert.That(errors, Has.Count.EqualTo(2));
            Assert.That(errors.Select(x => x.Severity), Is.All.EqualTo(TemplateSyntaxErrorSeverity.Warning));

            template.BindModel("Name", "foo");
            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Is.EqualTo("{{Name} foo }}"));
        }

        [Test]
        public void StructuralErrors_AreErrors_ExpressionAndTagIssues_AreWarnings()
        {
            var errors = Validate("{{#Items}}{{/}}{{/}}", "{?{(a}}x{{/}}", "{{Name}");

            Assert.That(errors.Select(x => (x.Tag, x.Severity)), Is.EqualTo(
            [
                ("{{/}}", TemplateSyntaxErrorSeverity.Error),
                ("{?{(a}}", TemplateSyntaxErrorSeverity.Warning),
                ("{{Name}", TemplateSyntaxErrorSeverity.Warning),
            ]));
        }
    }
}
