using System.Globalization;
using System.Net;

namespace DocxTemplater.Text.Test
{
    /// <summary>
    /// Templates as an HTML editor stores them, rendered with a preprocessor as an application would register one.
    /// The cases are values of <c>editor.value()</c> captured from Kendo UI 2026.3.
    /// </summary>
    internal class TagContentPreprocessorTest
    {
        /// <summary>Decodes character references and replaces typographic quotes and non-breaking spaces.</summary>
        private sealed class HtmlEditorPreprocessor : ITagContentPreprocessor
        {
            public string Process(string content, TagContentKind kind)
            {
                var decoded = WebUtility.HtmlDecode(content);
                var result = new char[decoded.Length];
                for (int i = 0; i < decoded.Length; i++)
                {
                    result[i] = decoded[i] switch
                    {
                        '‘' or '’' or '‚' or '‹' or '›' => '\'',
                        '“' or '”' or '„' or '«' or '»' => '"',
                        ' ' => ' ',
                        var c => c
                    };
                }
                return new string(result);
            }
        }

        private sealed class RecordingPreprocessor : ITagContentPreprocessor
        {
            public List<(string Content, TagContentKind Kind)> Calls { get; } = [];

            public string Process(string content, TagContentKind kind)
            {
                Calls.Add((content, kind));
                return content;
            }
        }

        private sealed class ReplacingPreprocessor(string oldValue, string newValue) : ITagContentPreprocessor
        {
            public string Process(string content, TagContentKind kind)
            {
                return content.Replace(oldValue, newValue, StringComparison.Ordinal);
            }
        }

        private sealed class Item
        {
            public string Name { get; set; }
            public decimal Price { get; set; }
        }

        private sealed class Recipient
        {
            public string FirstName { get; set; } = "Anna";
            public string LastName { get; set; } = "Müller & <Co>";
            public string Nothing { get; set; }
            public int Age { get; set; } = 30;
            public string City { get; set; } = "Bern";
            public bool IsVip { get; set; } = true;
            public int Count { get; set; } = 1;
            public int Max { get; set; } = 5;
            public int Id { get; set; } = 42;
        }

        private sealed class Mailing
        {
            public string UnsubscribeUrl { get; set; } = "https://orphy.ch/u/abc?x=1&y=2";
            public DateTime SentAt { get; set; } = new(2026, 11, 7);
            public List<Item> Items { get; set; } = [new Item { Name = "Kurs *A*", Price = 120.5m }, new Item { Name = "Kurs_B", Price = 80m }];
        }

        /// <param name="InvalidTags">The tags reported as <see cref="TemplateErrorCode.InvalidTag"/> and left as they are.</param>
        public sealed record EditorCase(string Name, string Html, string Expected, params string[] InvalidTags)
        {
            public override string ToString()
            {
                return Name;
            }
        }

        private const string PastedFromWord =
            "<span lang=\"DE-CH\" style=\"font-family:Calibri, sans-serif;\">Guten Tag {{recipient.</span><span lang=\"DE-CH\" style=\"font-family:Calibri, sans-serif;\">FirstName}}</span><span lang=\"DE-CH\">, Link: <a href=\"https://example.ch/x?id=%7B%7Brecipient.Id%7D%7D\">hier</a> und {{mailing.SentAt}:f(&lsquo;dd.MM.yyyy&rsquo;)} {?{recipient.City == &bdquo;Bern&ldquo;}}B{{/}}</span>";

        public static IEnumerable<EditorCase> EditorCases =>
        [
            new("character references in an expression",
                "Hallo {{recipient.FirstName}} {{recipient.LastName}}, Alter {{(recipient.Age &gt; 18 &amp;&amp; recipient.City == \"Bern\" ? \"ja\" : \"nein\")}} Datum {{mailing.SentAt}:f('dd.MM.yyyy')}",
                "Hallo Anna Müller &amp; &lt;Co&gt;, Alter ja Datum 07.11.2026"),
            new("less than without spaces",
                "<p>{?{recipient.Count&lt;recipient.Max}}klein{{/}}</p>",
                "<p>klein</p>"),
            new("typographic quotes in formatter arguments and conditions",
                "Guten Tag {{recipient.FirstName}}, {{mailing.SentAt}:f(&lsquo;dd.MM.yyyy&rsquo;)} {?{recipient.City == &bdquo;Bern&ldquo;}}B{{/}}",
                "Guten Tag Anna, 07.11.2026 B"),
            new("markup inside a tag",
                "<p>Hallo {{recipient.<strong>Last</strong>Name}}</p>",
                "<p>Hallo {{recipient.<strong>Last</strong>Name}}</p>",
                "{{recipient.<strong>Last</strong>Name}}"),
        ];

        private static TextTemplateSettings Plain(BindingErrorHandling errorHandling = BindingErrorHandling.ThrowException)
        {
            return new TextTemplateSettings
            {
                Escaping = TextEscaping.Html,
                Culture = new CultureInfo("de-CH"),
                UiCulture = CultureInfo.InvariantCulture,
                BindingErrorHandling = errorHandling
            };
        }

        private static TextTemplateSettings Settings(BindingErrorHandling errorHandling = BindingErrorHandling.ThrowException)
        {
            var settings = Plain(errorHandling);
            settings.RegisterTagContentPreprocessor(new HtmlEditorPreprocessor());
            return settings;
        }

        private static TextTemplate Template(string html, TextTemplateSettings settings = null)
        {
            var template = new TextTemplate(html, settings ?? Settings());
            template.BindModel("recipient", new Recipient());
            template.BindModel("mailing", new Mailing());
            return template;
        }

        [TestCaseSource(nameof(EditorCases))]
        public void EditorValueIsRendered(EditorCase editorCase)
        {
            var template = Template(editorCase.Html);
            var result = template.Process();
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo(editorCase.Expected));
                Assert.That(result.Errors, Is.Empty);
                Assert.That(template.ValidateTemplateSyntax().Select(x => (x.ErrorCode, x.Tag)), Is.EqualTo(editorCase.InvalidTags.Select(x => (TemplateErrorCode.InvalidTag, x))));
            });
        }

        [Test]
        public void WithoutPreprocessorTheTagsAreParsedAsWritten()
        {
            const string html = "<p>{?{recipient.Count &lt; recipient.Max}}klein{{/}}</p>";
            Assert.That(() => Template(html, Plain()).Process(), Throws.InstanceOf<TemplateException>());
        }

        [Test]
        public void OnlyTheContentOfRecognizedTagsIsPassed()
        {
            var recorder = new RecordingPreprocessor();
            var settings = Plain();
            settings.RegisterTagContentPreprocessor(recorder);
            new TextTemplate(
                "a &amp; b {{recipient.FirstName}} {{recipient.<b>Last</b>Name}} {?{recipient.Age &gt; 1}}y{{/}} " +
                "{{recipient.Nothing}:default('M &amp; Co, AG')} {{(recipient.Age)}:f(n0)} {{#recipient.Name}}{{/}} " +
                "{{#switch: recipient.City}}{{#case: 'Bern'}}b{{/}}{{/}}", settings).ValidateTemplateSyntax();
            (string, TagContentKind)[] expected =
            [
                ("recipient.Age &gt; 1", TagContentKind.Expression),
                ("'M &amp; Co, AG'", TagContentKind.FormatterArguments),
                ("(recipient.Age)", TagContentKind.Expression),
                ("n0", TagContentKind.FormatterArguments),
                (" recipient.City", TagContentKind.Expression),
                (" 'Bern'", TagContentKind.Expression),
            ];
            Assert.That(recorder.Calls, Is.EqualTo(expected));
        }

        [Test]
        public void FormatterArgumentsAreSplitAfterPreprocessing()
        {
            var template = Template("{{recipient.Nothing}:default(&lsquo;Smith, Jones&rsquo;)}|{{recipient.Nothing}:default('M&uuml;ller &amp; Co')}|{{recipient.FirstName}:default(&#39;x&#39;)}", Settings(BindingErrorHandling.HighlightErrorsInDocument));
            Assert.That(template.Process().Text, Is.EqualTo("Smith, Jones|Müller &amp; Co|Anna"));
        }

        [Test]
        public void PreprocessorsAreAppliedInRegistrationOrder()
        {
            var settings = Plain();
            settings.RegisterTagContentPreprocessor(new ReplacingPreprocessor("A", "B"));
            settings.RegisterTagContentPreprocessor(new ReplacingPreprocessor("B", "C"));
            Assert.That(new TextTemplate("{{('A')}}", settings).Process().Text, Is.EqualTo("C"));
        }

        [Test]
        public void ErrorPositionsReferToTheTemplateAsWritten()
        {
            const string html = "<p>&Uuml;ber {{recipient.<b>Miss</b>ing}} und {{#mailing.Items}}{?{.Price &gt; 100}}x</p>";
            var syntax = Template(html, Settings(BindingErrorHandling.HighlightErrorsInDocument)).ValidateTemplateSyntax();
            const string binding = "<p>&Uuml;ber {?{recipient.Missing &gt; 1}}x{{/}}</p>";
            var missing = Template(binding, Settings(BindingErrorHandling.HighlightErrorsInDocument)).Process().Errors.Single();
            Assert.Multiple(() =>
            {
                Assert.That(syntax.Select(x => x.Tag), Is.EquivalentTo(["{{recipient.<b>Miss</b>ing}}", "{{#mailing.Items}}", "{?{.Price &gt; 100}}"]));
                foreach (var error in syntax)
                {
                    Assert.That(html.Substring(error.Index, error.Length), Is.EqualTo(error.Tag));
                }
                Assert.That(missing.Tag, Is.EqualTo("{?{recipient.Missing &gt; 1}}"));
                Assert.That(binding.Substring(missing.Index, missing.Length), Is.EqualTo(missing.Tag));
            });
        }

        [Test]
        public void SchemaUsesThePreprocessedExpressions()
        {
            var schema = Template(PastedFromWord).GetTemplateSchema();
            Assert.Multiple(() =>
            {
                Assert.That(schema.Roots.Keys, Is.EquivalentTo(["recipient", "mailing"]));
                Assert.That(schema.Roots["recipient"].Properties.Keys, Is.EquivalentTo(["City"]));
                Assert.That(schema.Roots["mailing"].Properties.Keys, Is.EquivalentTo(["SentAt"]));
            });
        }

        [Test]
        public void TextOutsideOfTagsIsNotPreprocessed()
        {
            const string html = "<p title=\"a &amp; b\">Tom &amp; Jerry &lt;3 &bdquo;quoted&ldquo;&nbsp;{ x } {{recipient.FirstName}}</p>";
            Assert.That(Template(html).Process().Text, Is.EqualTo(html.Replace("{{recipient.FirstName}}", "Anna", StringComparison.Ordinal)));
        }
    }
}
