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
            var textTemplate = new TextTemplate(template, settings ?? Plain());
            textTemplate.BindModel("ds", model ?? new Person());
            return textTemplate.Process();
        }

        [Test]
        public void VariablesPathsAndImplicitPrefix()
        {
            Assert.That(Render("Hello {{ds.Name}}, {{Name}}, {{DS.NAME}}, {{ds.Items.Count}}"), Is.EqualTo("Hello Ada, Ada, Ada, 3"));
        }

        [Test]
        public void NullValueIsEmpty()
        {
            Assert.That(Render("[{{ds.Nothing}}]"), Is.EqualTo("[]"));
        }

        [Test]
        public void FormattersUseTheCulture()
        {
            var settings = Plain();
            settings.Culture = new CultureInfo("de-DE");
            Assert.That(Render("{{ds.Date}:f('dd.MM.yyyy')}|{{ds.Price}:format(N2)}|{{ds.Name}:toupper()}|{{ds.Name}:ToLower()}|{{ds.Shout}}", settings),
                Is.EqualTo("20.10.2026|1.234,50|ADA|ada|QUIET"));
        }

        [Test]
        public void ValueWithoutFormatterUsesTheCulture()
        {
            var settings = Plain();
            settings.Culture = new CultureInfo("de-DE");
            Assert.That(Render("{{ds.Price}}", settings), Is.EqualTo("1234,5"));
        }

        [Test]
        public void Expressions()
        {
            Assert.That(Render("{{(ds.Price * 2)}:f('0.00')}|{{(ds.Name.Length)}}|{{(ds.Nothing ?? \"none\")}}|{{(ds.Items[1].Name)}}"), Is.EqualTo("2469.00|3|none|b"));
        }

        [Test]
        public void LoopWithIndexLengthAndSeparator()
        {
            Assert.That(Render("{{#ds.Items}}{{ds.Items._Idx}}/{{Items._Length}}:{{.Name}}{{:s:}}, {{/ds.Items}}"), Is.EqualTo("1/3:a, 2/3:b, 3/3:c"));
        }

        [Test]
        public void SeparatorCanUseTheItem()
        {
            Assert.That(Render("{{#ds.Items}}{{.Name}}{{:s:}}({{.Name}}){{/ds.Items}}"), Is.EqualTo("a(a)b(b)c"));
        }

        [Test]
        public void NestedLoopsWithParentScope()
        {
            Assert.That(Render("{{#ds.Orders}}{{.Id}}:{{#.Lines}}{{..Id}}.{{.Name}}{{:s:}}+{{/.Lines}};{{/ds.Orders}}"), Is.EqualTo("7:7.x+7.y;8:8.z;"));
        }

        [Test]
        public void ConditionWithElse()
        {
            Assert.That(Render("{?{ds.Count > 1}}many{{:}}one{{/}}|{?{ds.Count > 5}}many{{else}}one{{/}}|{?{ds.Count > 5}}many{{/}}."), Is.EqualTo("many|one|."));
        }

        [Test]
        public void SwitchWithDefault()
        {
            const string template = "{{#switch: ds.Kind}}{{#case: 'A'}}alpha{{/}}{{#case: 'B'}}beta{{/}}{{#default}}other{{/}}{{/}}";
            Assert.Multiple(() =>
            {
                Assert.That(Render(template), Is.EqualTo("beta"));
                Assert.That(Render(template, model: new Person { Kind = "A" }), Is.EqualTo("alpha"));
                Assert.That(Render(template, model: new Person { Kind = "Z" }), Is.EqualTo("other"));
                Assert.That(Render("{{#s: ds.Count}}{{#c: 3}}three{{#c: 4}}four{{#d}}none{{/}}{{/}}"), Is.EqualTo("three"));
            });
        }

        [Test]
        public void SwitchInsideLoopIsEvaluatedPerItem()
        {
            Assert.That(Render("{{#ds.Items}}{{#switch: .Name}}{{#case: 'b'}}B{{/}}{{#default}}-{{/}}{{/}}{{/ds.Items}}"), Is.EqualTo("-B-"));
        }

        [Test]
        public void RangeLoops()
        {
            Assert.That(Render("{{@i:3}}{{i}}{{/}}|{{@ds.Count}}{{Index}}{{/}}|{{@n:ds.Items}}{{n}}{{/}}"), Is.EqualTo("012|012|012"));
        }

        [Test]
        public void IgnoreBlockAndBreak()
        {
            Assert.Multiple(() =>
            {
                Assert.That(Render("{{:ignore}}{{not.a.var}} {{#x}}{{/:ignore}}|a{{:break}}b"), Is.EqualTo("{{not.a.var}} {{#x}}|a\nb"));
                Assert.That(Render("a{{:break}}b", Html()), Is.EqualTo("a<br/>b"));
            });
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
            Assert.Multiple(() =>
            {
                Assert.That(template.Settings.Escaping, Is.EqualTo(TextEscaping.Html));
                Assert.That(template.Process(), Is.EqualTo("&lt;b&gt;bold&lt;/b&gt; &amp; more"));
            });
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
        public void WithoutIgnoreLineBreaksAroundTagsNothingIsTrimmed()
        {
            Assert.That(Render("{{#ds.Items}}\n{{.Name}}\n{{/ds.Items}}"), Is.EqualTo("\na\n\nb\n\nc\n"));
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

        [Test]
        public void BindingErrorsAreThrownForBlocks()
        {
            Assert.Multiple(() =>
            {
                // the core throws OpenXmlTemplateException, the text renderer TemplateException - both are a TemplateException
                Assert.That(Assert.Catch<TemplateException>(() => Render("{{#ds.Missing}}x{{/}}")).ErrorCode, Is.EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
                Assert.That(Assert.Catch<TemplateException>(() => Render("{{#ds.Count}}x{{/}}")).ErrorCode, Is.EqualTo(TemplateErrorCode.NotEnumerable));
                Assert.That(Assert.Catch<TemplateException>(() => Render("{{@i:ds.Name}}x{{/}}")).ErrorCode, Is.EqualTo(TemplateErrorCode.RangeCountNotNumeric));
                Assert.That(Assert.Catch<TemplateException>(() => Render("{?{ds.Missing > 1}}x{{/}}")).ErrorCode, Is.EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
                Assert.That(((TemplateException)Assert.Catch<TemplateException>(() => Render("{{ds.Date}:f()}")).InnerException).ErrorCode, Is.EqualTo(TemplateErrorCode.FormatPatternRequiresArgument));
                // a case formatter on a non-string is ignored, as in a Word template
                Assert.That(Render("{{ds.Count}:toupper()}"), Is.EqualTo("3"));
                // a string is enumerable - its characters are iterated, as in a Word template
                Assert.That(Render("{{#ds.Name}}{{.}}-{{/}}"), Is.EqualTo("A-d-a-"));
            });
        }

        [Test]
        public void SkipBindingAndRemoveContent()
        {
            var settings = Plain(BindingErrorHandling.SkipBindingAndRemoveContent);
            Assert.That(Render("A{{ds.Missing}}B{{#ds.Missing}}x{{/}}C{?{ds.Missing > 1}}yes{{:}}no{{/}}D{{@i:ds.Missing}}x{{/}}E", settings), Is.EqualTo("ABCnoDE"));
        }

        [Test]
        public void HighlightErrorsInPlainText()
        {
            var template = new TextTemplate("Hello {{ds.Missing}} and {{ds.Name}}", Plain(BindingErrorHandling.HighlightErrorsInDocument));
            template.BindModel("ds", new Person());
            var result = template.Process();
            var expectedError = $"Property 'Missing' not found in 'ds.Missing' of type '{typeof(Person)}'";
            Assert.Multiple(() =>
            {
                Assert.That(result, Is.EqualTo(expectedError + "\nHello {{ds.Missing}} and Ada"));
                Assert.That(template.Errors, Is.EqualTo([expectedError]));
            });
        }

        [Test]
        public void HighlightErrorsInHtml()
        {
            var template = new TextTemplate("<p>Hello {{ds.Missing}}</p>", Html(BindingErrorHandling.HighlightErrorsInDocument));
            template.BindModel("ds", new Person());
            var result = template.Process();
            var expectedError = $"Property 'Missing' not found in 'ds.Missing' of type '{typeof(Person)}'";
            Assert.Multiple(() =>
            {
                Assert.That(result, Does.StartWith("<div style=\"color:#ff0000;font-weight:bold\">Property &#39;Missing&#39; not found"));
                Assert.That(result, Does.EndWith("</div><p>Hello <span style=\"background-color:#ff0000;color:#000000;font-weight:bold\">{{ds.Missing}}</span></p>"));
                Assert.That(template.Errors, Is.EqualTo([expectedError]));
            });
        }

        [Test]
        public void HighlightListsEachErrorOnce()
        {
            var template = new TextTemplate("{{#ds.Items}}{{.Missing}}{{/ds.Items}}", Plain(BindingErrorHandling.HighlightErrorsInDocument));
            template.BindModel("ds", new Person());
            var result = template.Process();
            Assert.Multiple(() =>
            {
                Assert.That(template.Errors, Has.Count.EqualTo(3));
                Assert.That(result.Split('\n'), Has.Length.EqualTo(2));
            });
        }

        [Test]
        public void SyntaxErrorsAreReportedAndMakeProcessThrow()
        {
            var template = new TextTemplate("{{#ds.Items}}{{.Name}}", Plain());
            template.BindModel("ds", new Person());
            var errors = template.ValidateTemplateSyntax();
            Assert.Multiple(() =>
            {
                Assert.That(errors, Has.Count.EqualTo(1));
                Assert.That(errors[0].ErrorCode, Is.EqualTo(TemplateErrorCode.BlockNotClosed));
                Assert.That(errors[0].Part, Is.EqualTo("Text"));
                Assert.That(errors[0].Index, Is.EqualTo(0));
                Assert.That(errors[0].Length, Is.EqualTo("{{#ds.Items}}".Length));
                var e = Assert.Throws<TemplateException>(() => template.Process());
                Assert.That(e.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
                Assert.That(e.Message, Does.Contain("Text: '{{#ds.Items}}' is not closed"));
                Assert.That(Assert.Throws<TemplateException>(() => template.GetTemplateSchema()).ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
            });
        }

        [Test]
        public void SyntaxErrorsAreHighlightedInsteadOfRendered()
        {
            var template = new TextTemplate("<p>{{ds.Name}} {{#ds.Items}}{{.Name}}</p>", Html(BindingErrorHandling.HighlightErrorsInDocument));
            template.BindModel("ds", new Person());
            var result = template.Process();
            Assert.Multiple(() =>
            {
                Assert.That(result, Does.EndWith("</div><p>{{ds.Name}} <span style=\"background-color:#ff0000;color:#000000;font-weight:bold\">{{#ds.Items}}</span>{{.Name}}</p>"));
                Assert.That(template.Errors, Has.Count.EqualTo(1));
                Assert.That(template.Errors[0], Does.StartWith("Template syntax errors:"));
            });
        }

        [Test]
        public void WordOnlyConstructsAreSyntaxErrors()
        {
            var template = new TextTemplate("{{ds.Name}:img()}{{:PageBreak}}{{#ds.Items}:dyntable()}{{/}}{{ds.Name}:T()}{{ds.Name}:md()}{{:break}}{{ds.Name}:html()}", Plain());
            var errors = template.ValidateTemplateSyntax();
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
                Assert.That(Assert.Throws<TemplateException>(() => template.Process()).ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
            });
        }

        [Test]
        public void UnknownFormatterFallsBackToToString()
        {
            Assert.That(Render("{{ds.Name}:unknown()}"), Is.EqualTo("Ada"));
        }

        [Test]
        public void SchemaOfTheTemplate()
        {
            var template = new TextTemplate("{{ds.Name}}{{#ds.Items}}{{.Price}}{?{.Name == 'x'}}{{..Count}}{{/}}{{/ds.Items}}{{(ds.Date.Year)}}{{@i:ds.Count}}{{i}}{{/}}", Plain());
            var schema = template.GetTemplateSchema();
            var ds = schema.Roots["ds"];
            Assert.Multiple(() =>
            {
                Assert.That(ds.Properties.Keys, Is.EquivalentTo(["Name", "Items", "Date", "Count"]));
                Assert.That(ds.Properties["Items"].Kind, Is.EqualTo(Schema.TemplateNodeKind.Collection));
                Assert.That(ds.Properties["Items"].ItemSchema.Properties.Keys, Is.EquivalentTo(["Price", "Name"]));
                Assert.That(ds.Properties["Date"].Properties.Keys, Is.EquivalentTo(["Year"]));
                Assert.That(template.GetTemplateSchema(), Is.SameAs(schema));
            });
        }

        [Test]
        public void ProcessCanBeRepeated()
        {
            var template = new TextTemplate("{{#ds.Items}}{{.Name}}{{/ds.Items}}", Plain());
            template.BindModel("ds", new Person());
            Assert.Multiple(() =>
            {
                Assert.That(template.Process(), Is.EqualTo("abc"));
                Assert.That(template.Process(), Is.EqualTo("abc"));
                Assert.That(template.Models.Keys, Is.EqualTo(["ds"]));
            });
        }

        private sealed class ReverseFormatter : ITextFormatter
        {
            public bool CanHandle(Type type, string prefix)
            {
                return type == typeof(string) && prefix.Equals("reverse", StringComparison.OrdinalIgnoreCase);
            }

            public TextValue ApplyFormat(ITextTemplateContext context, FormatterContext formatterContext)
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
            var template = new TextTemplate("{{ds.Html}:reverse()}|{{ds.Html}:reverse(raw)}", Html());
            template.BindModel("ds", new Person { Html = "<a>" });
            template.RegisterFormatter(new ReverseFormatter());
            Assert.That(template.Process(), Is.EqualTo("&gt;a&lt;|>a<"));
        }

        [Test]
        public void DictionaryModel()
        {
            var model = new Dictionary<string, object> { ["Name"] = "Bob", ["Tags"] = new[] { "x", "y" } };
            Assert.That(Render("{{ds.name}}:{{#ds.Tags}}{{.}}{{:s:}},{{/ds.Tags}}", model: model), Is.EqualTo("Bob:x,y"));
        }

        [Test]
        public void ErrorMessagesUseTheUiCulture()
        {
            var settings = Plain();
            settings.UiCulture = new CultureInfo("fr");
            settings.ErrorMessages = new TemplateErrorMessages { AutoLoadLanguagePackages = false }
                .AddLanguage(new CultureInfo("fr"), new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.BlockNotClosed] = "'{0}' n'est pas fermé" });
            var template = new TextTemplate("{{#ds.Items}}", settings);
            Assert.That(template.ValidateTemplateSyntax()[0].Message, Is.EqualTo("'{{#ds.Items}}' n'est pas fermé"));
        }
    }
}
