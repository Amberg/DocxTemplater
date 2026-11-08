using System.Net;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class TagContentPreprocessorTest
    {
        /// <summary>
        /// Decodes character references and replaces the typographic quotes of Word's autocorrect; formatter arguments
        /// are quoted with single quotes.
        /// </summary>
        private sealed class QuotesPreprocessor : ITagContentPreprocessor
        {
            public string Process(string content, TagContentKind kind)
            {
                var doubleQuote = kind == TagContentKind.FormatterArguments ? '\'' : '"';
                return WebUtility.HtmlDecode(content)
                    .Replace('‘', '\'').Replace('’', '\'')
                    .Replace('“', doubleQuote).Replace('”', doubleQuote).Replace('„', doubleQuote);
            }
        }

        private sealed class Model
        {
            public int Count { get; set; } = 3;
            public string City { get; set; } = "Bern";
            public string Nothing { get; set; }
            public DateTime Date { get; set; } = new(2026, 11, 7);
        }

        private static MemoryStream CreateDocument()
        {
            var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    // a tag split over runs, as Word stores it after editing
                    new Paragraph(new Run(new Text("{?{ds.Count ")), new Run(new Text("&gt; 1}}big{{/}}"))),
                    new Paragraph(new Run(new Text("{{(ds.City == „Bern“ ? ‘B’ : ‘X’)}}"))),
                    new Paragraph(new Run(new Text("{{ds.Nothing}:default(“n/a, none”)}|{{ds.Date}:f(‘dd.MM.yyyy’)}")))));
                wpDocument.Save();
            }
            memStream.Position = 0;
            return memStream;
        }

        private static readonly string[] Expected = ["big", "B", "n/a, none|07.11.2026"];

        private static string[] Paragraphs(Stream result)
        {
            using var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body.Elements<Paragraph>().Select(x => x.InnerText).ToArray();
        }

        private static ProcessSettings WithPreprocessor(BindingErrorHandling errorHandling = BindingErrorHandling.ThrowException)
        {
            var settings = new ProcessSettings { BindingErrorHandling = errorHandling };
            settings.RegisterTagContentPreprocessor(new QuotesPreprocessor());
            return settings;
        }

        [Test]
        public void ExpressionsAndFormatterArgumentsArePreprocessed()
        {
            using var template = new DocxTemplate(CreateDocument(), WithPreprocessor());
            template.BindModel("ds", new Model());
            var result = template.Process();
            template.Validate();
            Assert.That(Paragraphs(result), Is.EqualTo(Expected));
        }

        [Test]
        public void SchemaAndValidationUseThePreprocessedTags()
        {
            using var template = new DocxTemplate(CreateDocument(), WithPreprocessor());
            template.BindModel("ds", new Model());
            Assert.Multiple(() =>
            {
                Assert.That(template.ValidateTemplateSyntax(), Is.Empty);
                Assert.That(template.GetTemplateSchema().Roots["ds"].Properties.Keys, Is.EquivalentTo(["Count", "City", "Nothing", "Date"]));
                Assert.That(Paragraphs(template.Process()), Is.EqualTo(Expected));
            });
        }

        [Test]
        public void WithoutPreprocessorTheTagsAreParsedAsWritten()
        {
            using var template = new DocxTemplate(CreateDocument());
            template.BindModel("ds", new Model());
            Assert.That(template.Process, Throws.InstanceOf<OpenXmlTemplateException>());
        }

        [Test]
        public void EachTemplateWorksWithACopyOfTheSettings()
        {
            static string[] Process(DocxTemplate template)
            {
                template.BindModel("ds", new Model());
                return Paragraphs(template.Process());
            }

            var settings = new ProcessSettings();
            using var before = new DocxTemplate(CreateDocument(), settings);
            using var ownChange = new DocxTemplate(CreateDocument(), settings);
            ownChange.Settings.RegisterTagContentPreprocessor(new QuotesPreprocessor());
            using var notAffected = new DocxTemplate(CreateDocument(), settings);
            settings.RegisterTagContentPreprocessor(new QuotesPreprocessor());
            using var after = new DocxTemplate(CreateDocument(), settings);
            Assert.Multiple(() =>
            {
                Assert.That(() => Process(before), Throws.InstanceOf<OpenXmlTemplateException>());
                Assert.That(Process(ownChange), Is.EqualTo(Expected));
                Assert.That(() => Process(notAffected), Throws.InstanceOf<OpenXmlTemplateException>());
                Assert.That(Process(after), Is.EqualTo(Expected));
            });
        }

        [Test]
        public void TemplatesProcessedInParallelShareTheSettings()
        {
            var settings = WithPreprocessor();
            var models = Enumerable.Range(0, 40).Select(i => new Model { Count = i % 3, City = i % 2 == 0 ? "Bern" : "Basel" }).ToList();
            string[] Render(Model model)
            {
                using var template = new DocxTemplate(CreateDocument(), settings);
                template.BindModel("ds", model);
                return Paragraphs(template.Process());
            }

            var results = new string[models.Count][];
            Parallel.For(0, models.Count, i => results[i] = Render(models[i]));
            for (int i = 0; i < models.Count; i++)
            {
                var city = models[i].City == "Bern" ? "B" : "X";
                // the paragraph of a false condition is removed
                string[] expected = models[i].Count > 1 ? ["big", city, Expected[2]] : [city, Expected[2]];
                Assert.That(results[i], Is.EqualTo(expected), $"model {i}");
            }
        }
    }
}
