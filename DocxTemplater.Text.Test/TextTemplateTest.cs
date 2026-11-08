using System.Globalization;
using DocxTemplater.Formatter;
using DocxTemplater.Model;

namespace DocxTemplater.Text.Test
{
    internal class TextTemplateTest
    {
        private sealed class Item
        {
            public string Name { get; set; }
            public decimal Price { get; set; }
        }

        private sealed class Order
        {
            public int Id { get; set; }
            public List<Item> Lines { get; set; }
        }

        private sealed class Person
        {
            public string Name { get; set; } = "Ada";
            public string Html { get; set; } = "<b>bold</b> & more";
            public string Nothing { get; set; }
            public DateTime Date { get; set; } = new(2026, 10, 20);
            public decimal Price { get; set; } = 1234.5m;
            public int Count { get; set; } = 3;
            public string Kind { get; set; } = "B";
            public List<Item> Items { get; set; } = [new Item { Name = "a", Price = 1 }, new Item { Name = "b", Price = 2 }, new Item { Name = "c", Price = 3 }];
            public List<Order> Orders { get; set; } = [new Order { Id = 7, Lines = [new Item { Name = "x" }, new Item { Name = "y" }] }, new Order { Id = 8, Lines = [new Item { Name = "z" }] }];

            [ModelProperty(DefaultFormatter = "toupper")]
            public string Shout { get; set; } = "quiet";
        }

        private static TextTemplateSettings Plain(BindingErrorHandling errorHandling = BindingErrorHandling.ThrowException)
        {
            return new TextTemplateSettings
            {
                Escaping = TextEscaping.None,
                Culture = CultureInfo.InvariantCulture,
                UiCulture = CultureInfo.InvariantCulture,
                BindingErrorHandling = errorHandling
            };
        }

        private static TextTemplateSettings Html(BindingErrorHandling errorHandling = BindingErrorHandling.ThrowException)
        {
            var settings = Plain(errorHandling);
            settings.Escaping = TextEscaping.Html;
            return settings;
        }

        private static string Render(string template, TextTemplateSettings settings = null, object model = null)
        {
            return Process(template, settings, model).Text;
        }

        private static TextRenderResult Process(string template, TextTemplateSettings settings = null, object model = null)
        {
            var textTemplate = Template(template, settings);
            textTemplate.BindModel("ds", model ?? new Person());
            return textTemplate.Process();
        }

        private static TextTemplate Template(string template, TextTemplateSettings settings = null)
        {
            return new TextTemplate(template, settings ?? Plain());
        }

        /// <summary>The error codes of an exception and its inner exceptions, outermost first.</summary>
        private static IEnumerable<TemplateErrorCode> ErrorCodes(Exception e)
        {
            for (; e != null; e = e.InnerException)
            {
                if (e is TemplateException templateException)
                {
                    yield return templateException.ErrorCode;
                }
            }
        }

        [Test]
        public void BreakInHtml()
        {
            Assert.That(Render("a{{:break}}b", Html()), Is.EqualTo("a<br/>b"));
        }

        [Test]
        public void HtmlEscapingEncodesValuesButNotTheTemplate()
        {
            Assert.That(Render("<p>{{ds.Html}}</p><p>{{ds.Html}:html()}</p>", Html()),
                Is.EqualTo("<p>&lt;b&gt;bold&lt;/b&gt; &amp; more</p><p><b>bold</b> & more</p>"));
        }

        [Test]
        public void HtmlEscapingTurnsLineBreaksIntoBr()
        {
            var model = new Person { Name = "line1\r\nline2\nline3" };
            Assert.Multiple(() =>
            {
                Assert.That(Render("{{ds.Name}}", Html(), model), Is.EqualTo("line1<br/>line2<br/>line3"));
                Assert.That(Render("{{ds.Name}}", Plain(), model), Is.EqualTo("line1\r\nline2\nline3"));
            });
        }

        [Test]
        public void EscapingDefaultsToHtml()
        {
            var template = new TextTemplate("{{ds.Html}}");
            template.BindModel("ds", new Person());
            Assert.That(template.Process().Text, Is.EqualTo("&lt;b&gt;bold&lt;/b&gt; &amp; more"));
        }

        [Test]
        public void IgnoreLineBreaksAroundTagsRemovesStandaloneTagLines()
        {
            var settings = Plain();
            settings.IgnoreLineBreaksAroundTags = true;
            const string template = "Items:\n{{#ds.Items}}\n  - {{.Name}}\n{{/ds.Items}}\n  {?{ds.Count > 1}}  \nmany\n{{:}}\none\n{{/}}\nEnd {{ds.Name}}\n{{ds.Name}}\n";
            Assert.That(Render(template, settings), Is.EqualTo("Items:\n  - a\n  - b\n  - c\nmany\nEnd Ada\nAda\n"));
        }

        [Test]
        public void IgnoreLineBreaksAroundTagsKeepsWindowsLineBreaksAndInlineTags()
        {
            var settings = Plain();
            settings.IgnoreLineBreaksAroundTags = true;
            Assert.Multiple(() =>
            {
                Assert.That(Render("A\r\n{{#ds.Items}}\r\n{{.Name}}\r\n{{/ds.Items}}\r\nB", settings), Is.EqualTo("A\r\na\r\nb\r\nc\r\nB"));
                Assert.That(Render("{{#ds.Items}}{{.Name}}{{/ds.Items}}\nB", settings), Is.EqualTo("abc\nB"));
                Assert.That(Render("{{#ds.Items}}{{/ds.Items}}{{#ds.Items}} {{/ds.Items}}\nB", settings), Is.EqualTo("   B"));
            });
        }

        [Test]
        public void BindingErrorThrowsWithContext()
        {
            var e = Assert.Throws<TemplateException>(() => Render("Hello {{ds.Missing}} there"));
            Assert.Multiple(() =>
            {
                Assert.That(e.ErrorCode, Is.EqualTo(TemplateErrorCode.PlaceholderNotReplaced));
                Assert.That(e.Arguments, Is.EqualTo(new object[] { "{{ds.Missing}}", "Hello ", " there" }));
                Assert.That(e.InnerException, Is.InstanceOf<TemplateException>().With.Property("ErrorCode").EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
                Assert.That(e.Message, Does.Contain("Hello  >> {{ds.Missing}} <<  there"));
            });
        }

        [TestCase("{{#ds.Missing}}x{{/}}", "{{#ds.Missing}}", TemplateErrorCode.PropertyNotFoundOnType)]
        [TestCase("{{#ds.Count}}x{{/}}", "{{#ds.Count}}", TemplateErrorCode.NotEnumerable)]
        [TestCase("{{@i:ds.Name}}x{{/}}", "{{@i:ds.Name}}", TemplateErrorCode.RangeCountNotNumeric)]
        [TestCase("{?{ds.Missing > 1}}x{{/}}", "{?{ds.Missing > 1}}", TemplateErrorCode.ErrorInCondition, TemplateErrorCode.PropertyNotFoundOnType)]
        [TestCase("{{#switch: ds.Missing}}{{#case: 1}}x{{/}}{{/}}", "{{#case: 1}}", TemplateErrorCode.ErrorInSwitchCase, TemplateErrorCode.PropertyNotFoundOnType)]
        [TestCase("{{ds.Date}:f()}", "{{ds.Date}:f()}", TemplateErrorCode.FormatPatternRequiresArgument)]
        [TestCase("{{ds.Count}:toupper()}", "{{ds.Count}:toupper()}", TemplateErrorCode.UnknownFormatter)]
        [TestCase("{{(ds.Items[5])}}", "{{(ds.Items[5])}}", TemplateErrorCode.None)]
        public void EveryTagIsReportedWithItsContextOnThrow(string tag, string failedTag, params TemplateErrorCode[] causes)
        {
            var e = Assert.Throws<TemplateException>(() => Render("before " + tag + " after"));
            Assert.Multiple(() =>
            {
                Assert.That(ErrorCodes(e), Is.EqualTo(causes.Prepend(TemplateErrorCode.PlaceholderNotReplaced)));
                Assert.That(e.Arguments[0], Is.EqualTo(failedTag));
                Assert.That(e.Message, Does.Contain(" >> " + failedTag + " << "));
            });
        }

        [Test]
        public void HighlightErrorsInHtml()
        {
            var result = Process("<p>Hello {{ds.Missing}}</p>", Html(BindingErrorHandling.HighlightErrorsInDocument));
            var expectedError = $"Property 'Missing' not found in 'ds.Missing' of type '{typeof(Person)}'";
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Does.StartWith("<div style=\"color:#ff0000;font-weight:bold\">Property &#39;Missing&#39; not found"));
                Assert.That(result.Text, Does.EndWith("</div><p>Hello <span style=\"background-color:#ff0000;color:#000000;font-weight:bold\">{{ds.Missing}}</span></p>"));
                Assert.That(result.Errors.Select(x => x.Exception.Message), Is.EqualTo([expectedError]));
            });
        }

        [Test]
        public void SyntaxErrorsMakeProcessThrow()
        {
            const string template = "{{#ds.Items}}{{.Name}}";
            var e = Assert.Throws<TemplateException>(() => Process(template));
            var highlighted = Template(template, Plain(BindingErrorHandling.HighlightErrorsInDocument));
            var errors = highlighted.ValidateTemplateSyntax();
            Assert.Multiple(() =>
            {
                Assert.That(e.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
                Assert.That(e.Message, Does.Contain("Text: '{{#ds.Items}}' is not closed"));
                Assert.That(((IReadOnlyList<TemplateSyntaxError>)e.Arguments.Single()).Single().ErrorCode, Is.EqualTo(TemplateErrorCode.BlockNotClosed));
                Assert.That(errors, Has.Count.EqualTo(1));
                Assert.That(errors[0].ErrorCode, Is.EqualTo(TemplateErrorCode.BlockNotClosed));
                Assert.That(errors[0].Part, Is.EqualTo("Text"));
                Assert.That(errors[0].Index, Is.EqualTo(0));
                Assert.That(errors[0].Length, Is.EqualTo("{{#ds.Items}}".Length));
                Assert.That(Assert.Throws<TemplateException>(() => highlighted.GetTemplateSchema()).ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
            });
        }

        [Test]
        public void SyntaxErrorsAreHighlightedInsteadOfRendered()
        {
            var result = Process("<p>{{ds.Name}} {{#ds.Items}}{{.Name}}</p>", Html(BindingErrorHandling.HighlightErrorsInDocument));
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Does.EndWith("</div><p>{{ds.Name}} <span style=\"background-color:#ff0000;color:#000000;font-weight:bold\">{{#ds.Items}}</span>{{.Name}}</p>"));
                Assert.That(result.Errors.Single().Exception.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
            });
        }

        [Test]
        public void WordOnlyConstructsAreSyntaxErrors()
        {
            const string template = "{{ds.Name}:img()}{{:PageBreak}}{{#ds.Items}:dyntable()}{{/}}{{ds.Name}:T()}{{ds.Name}:md()}{{:break}}{{ds.Name}:html()}";
            var errors = Template(template, Plain(BindingErrorHandling.HighlightErrorsInDocument)).ValidateTemplateSyntax();
            Assert.Multiple(() =>
            {
                Assert.That(errors.Select(x => x.ErrorCode), Is.EqualTo(new List<TemplateErrorCode>
                {
                    TemplateErrorCode.FormatterNotSupported, TemplateErrorCode.KeywordNotSupported, TemplateErrorCode.FormatterNotSupported,
                    TemplateErrorCode.FormatterNotSupported, TemplateErrorCode.FormatterNotSupported
                }));
                Assert.That(errors.All(x => x.Severity == TemplateSyntaxErrorSeverity.Error));
                Assert.That(errors[0].Arguments, Is.EqualTo(new object[] { "{{ds.Name}:img()}", "img" }));
                Assert.That(errors[0].Message, Is.EqualTo("'{{ds.Name}:img()}': the formatter 'img' is only available in Word documents"));
                Assert.That(errors[1].Message, Is.EqualTo("'{{:PageBreak}}': this keyword is only available in Word documents"));
                Assert.That(Assert.Throws<TemplateException>(() => Process(template)).ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
            });
        }

        [Test]
        public void SchemaOfTheTemplate()
        {
            var schema = Template("{{ds.Name}}{{#ds.Items}}{{.Price}}{?{.Name == 'x'}}{{..Count}}{{/}}{{/ds.Items}}{{(ds.Date.Year)}}{{@i:ds.Count}}{{i}}{{/}}").GetTemplateSchema();
            var ds = schema.Roots["ds"];
            Assert.Multiple(() =>
            {
                Assert.That(ds.Properties.Keys, Is.EquivalentTo(["Name", "Items", "Date", "Count"]));
                Assert.That(ds.Properties["Items"].Kind, Is.EqualTo(Schema.TemplateNodeKind.Collection));
                Assert.That(ds.Properties["Items"].ItemSchema.Properties.Keys, Is.EquivalentTo(["Price", "Name"]));
                Assert.That(ds.Properties["Date"].Properties.Keys, Is.EquivalentTo(["Year"]));
            });
        }

        private sealed class ReverseFormatter : ITextFormatter
        {
            public bool CanHandle(Type type, string prefix)
            {
                return type == typeof(string) && prefix.Equals("reverse", StringComparison.OrdinalIgnoreCase);
            }

            public TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext)
            {
                var chars = ((string)formatterContext.Value).ToCharArray();
                Array.Reverse(chars);
                var reversed = new string(chars);
                return formatterContext.Args.Length > 0 && formatterContext.Args[0] == "raw" ? TextValue.Raw(reversed) : new TextValue(reversed);
            }
        }

        [Test]
        public void CustomFormatter()
        {
            var settings = Html();
            settings.RegisterFormatter(new ReverseFormatter());
            var template = new TextTemplate("{{ds.Html}:reverse()}|{{ds.Html}:reverse(raw)}", settings);
            template.BindModel("ds", new Person { Html = "<a>" });
            Assert.That(template.Process().Text, Is.EqualTo("&gt;a&lt;|>a<"));
        }

        [Test]
        public void ErrorMessagesUseTheUiCulture()
        {
            var settings = Plain(BindingErrorHandling.HighlightErrorsInDocument);
            settings.UiCulture = new CultureInfo("fr");
            settings.ErrorMessages = new TemplateErrorMessages { AutoLoadLanguagePackages = false }
                .AddLanguage(new CultureInfo("fr"), new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.BlockNotClosed] = "'{0}' n'est pas fermé" });
            Assert.That(Template("{{#ds.Items}}", settings).ValidateTemplateSyntax()[0].Message, Is.EqualTo("'{{#ds.Items}}' n'est pas fermé"));
        }

        [Test]
        public void DefaultTextIsEncodedLikeAValue()
        {
            Assert.That(Render("{{ds.Nothing}:default('Smith & <Sons>')}|{{ds.Nothing}:default()}|", Html()), Is.EqualTo("Smith &amp; &lt;Sons&gt;||"));
        }

        [Test]
        public void HtmlEscapingEncodesOnlyWhatHtmlRequires()
        {
            var person = new Person { Name = "Müller & <Co> \"x\" 'y'\nzwei" };
            Assert.Multiple(() =>
            {
                Assert.That(Render("<p>{{ds.Name}}</p>", Html(), person), Is.EqualTo("<p>Müller &amp; &lt;Co&gt; &quot;x&quot; &#39;y&#39;<br/>zwei</p>"));
                Assert.That(Render("{{ds.Name}}", Plain(), person), Is.EqualTo(person.Name));
            });
        }
    }
}
