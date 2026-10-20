using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Text.Test
{
    /// <summary>
    /// Runs the same template cases through the Word renderer and the text renderer, so the two cannot drift
    /// apart. A line of the template is a paragraph of the Word document; the paragraphs of the result are joined
    /// with a line feed and compared with the text output modulo trailing line feeds.
    /// </summary>
    internal class ConformanceTest
    {
        public sealed class Item
        {
            public string Name { get; set; }
            public decimal Price { get; set; }
        }

        public sealed class Order
        {
            public int Id { get; set; }
            public List<Item> Lines { get; set; }
        }

        public sealed class Address
        {
            public string City { get; set; } = "Bern";
        }

        public sealed class Person
        {
            public string Name { get; set; } = "Ada";
            public string Nothing { get; set; }
            public Address Address { get; set; } = new();
            public DateTime Date { get; set; } = new(2026, 10, 20);
            public decimal Price { get; set; } = 1234.5m;
            public int Count { get; set; } = 3;
            public string Kind { get; set; } = "B";
            public List<Item> Items { get; set; } = [new Item { Name = "a", Price = 1 }, new Item { Name = "b", Price = 2 }, new Item { Name = "c", Price = 3 }];
            public List<Order> Orders { get; set; } = [new Order { Id = 7, Lines = [new Item { Name = "x" }, new Item { Name = "y" }] }, new Order { Id = 8, Lines = [new Item { Name = "z" }] }];

            [Model.ModelProperty(DefaultFormatter = "toupper")]
            public string Shout { get; set; } = "quiet";
        }

        public sealed record Case(string Name, string Template, string Expected)
        {
            public Func<object> Model { get; init; } = () => new Person();
            public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;
            public BindingErrorHandling ErrorHandling { get; init; } = BindingErrorHandling.ThrowException;
            public bool IgnoreLineBreaksAroundTags { get; init; }

            /// <summary>The error <c>Process</c> is expected to throw; <see cref="TemplateErrorCode.None"/> if it is expected to render.</summary>
            public TemplateErrorCode ExpectedError { get; init; }

            public override string ToString()
            {
                return Name;
            }
        }

        private static readonly CultureInfo German = new("de-DE");

        public static IEnumerable<Case> Cases =>
        [
            new("variables and paths", "Hello {{ds.Name}} from {{DS.address.city}}, {{Name}}", "Hello Ada from Bern, Ada"),
            new("null value", "[{{ds.Nothing}}]", "[]"),
            new("formatters with culture", "{{ds.Date}:f('yyyy-MM-dd')}|{{ds.Price}:F(N2)}|{{ds.Name}:toupper()}|{{ds.Name}:ToLower()}|{{ds.Shout}}", "2026-10-20|1.234,50|ADA|ada|QUIET") { Culture = German },
            new("value in culture", "{{ds.Price}}|{{ds.Date}}", "1234,5|20.10.2026 00:00:00") { Culture = German },
            new("expressions", "{{(ds.Price * 2)}:f('0.00')}|{{(ds.Name.Length)}}|{{(ds.Nothing ?? \"none\")}}|{{(ds.Items[1].Name)}}|{{(ds.Items.Count > 2 ? \"many\" : \"few\")}}", "2469.00|3|none|b|many"),
            new("loop with index, length and separator", "{{#ds.Items}}{{ds.Items._Idx}}/{{Items._Length}}:{{.Name}}{{:s:}}, {{/ds.Items}}", "1/3:a, 2/3:b, 3/3:c"),
            new("separator uses the item", "{{#ds.Items}}{{.Name}}{{:s:}}({{.Name}}){{/ds.Items}}", "a(a)b(b)c"),
            new("nested loops with parent scope", "{{#ds.Orders}}{{.Id}}:{{#.Lines}}{{..Id}}.{{.Name}}{{:s:}}+{{/.Lines}};{{/ds.Orders}}", "7:7.x+7.y;8:8.z;"),
            new("loop over a string", "{{#ds.Name}}{{.}}-{{/}}", "A-d-a-"),
            new("empty loop", "[{{#ds.Items}}{{.Name}}{{:s:}},{{/}}]", "[]") { Model = () => new Person { Items = [] } },
            new("condition with else", "{?{ds.Count > 1}}many{{:}}one{{/}}|{?{ds.Count > 5}}many{{else}}one{{/}}|{?{ds.Count > 5}}x{{/}}.", "many|one|."),
            new("condition inside loop", "{{#ds.Items}}{?{.Price > 1}}{{.Name}}{{:}}-{{/}}{{/ds.Items}}", "-bc"),
            new("switch", "{{#switch: ds.Kind}}{{#case: 'A'}}alpha{{/}}{{#case: 'B'}}beta{{/}}{{#default}}other{{/}}{{/}}", "beta"),
            new("switch default", "{{#switch: ds.Kind}}{{#case: 'A'}}alpha{{/}}{{#default}}other{{/}}{{/}}", "other") { Model = () => new Person { Kind = "Z" } },
            new("switch short form with implicit closing", "{{#s: ds.Count}}{{#c: 3}}three{{#c: 4}}four{{#d}}none{{/}}{{/}}", "three"),
            new("switch inside loop", "{{#ds.Items}}{{#switch: .Name}}{{#case: 'b'}}B{{/}}{{#default}}-{{/}}{{/}}{{/ds.Items}}", "-B-"),
            new("range loops", "{{@i:3}}{{i}}{{/}}|{{@ds.Count}}{{Index}}{{/}}|{{@n:ds.Items}}{{n}}{{/}}", "012|012|012"),
            new("ignore block", "{{:ignore}}{{not.a.var}} {{#x}}{{/:ignore}}", "{{not.a.var}} {{#x}}"),
            new("break", "a{{:break}}b", "a\nb"),
            new("multi-line value", "{{ds.Name}}", "l1\nl2") { Model = () => new Person { Name = "l1\nl2" } },
            new("loop over lines", "Items:\n{{#ds.Items}}\n- {{.Name}}\n{{/ds.Items}}\nEnd", "Items:\n- a\n- b\n- c\nEnd") { IgnoreLineBreaksAroundTags = true },
            new("condition over lines", "{?{ds.Count > 1}}\nmany\n{{:}}\none\n{{/}}", "many") { IgnoreLineBreaksAroundTags = true },
            new("dictionary model", "{{ds.name}}:{{#ds.Tags}}{{.}}{{:s:}},{{/ds.Tags}}", "Bob:x,y") { Model = () => new Dictionary<string, object> { ["Name"] = "Bob", ["Tags"] = new[] { "x", "y" } } },
            new("skip binding errors", "A{{ds.Missing}}B{{#ds.Missing}}x{{/}}C{?{ds.Missing > 1}}yes{{:}}no{{/}}D{{@i:ds.Missing}}x{{/}}E", "ABCnoDE") { ErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent },
            new("highlight binding errors", "Hello {{ds.Missing}} and {{ds.Name}}", $"Property 'Missing' not found in 'ds.Missing' of type '{typeof(Person)}'\nHello {{{{ds.Missing}}}} and Ada") { ErrorHandling = BindingErrorHandling.HighlightErrorsInDocument },
            new("highlight lists each error once", "{{#ds.Items}}{{.Missing}}{{/ds.Items}}", $"Property 'Missing' not found in 'Missing' of type '{typeof(Item)}'\n{{{{.Missing}}}}{{{{.Missing}}}}{{{{.Missing}}}}") { ErrorHandling = BindingErrorHandling.HighlightErrorsInDocument },
            new("throw binding error", "Hello {{ds.Missing}}", null) { ExpectedError = TemplateErrorCode.PlaceholderNotReplaced },
            new("throw not enumerable", "{{#ds.Count}}x{{/}}", null) { ExpectedError = TemplateErrorCode.NotEnumerable },
            new("throw range count", "{{@i:ds.Name}}x{{/}}", null) { ExpectedError = TemplateErrorCode.RangeCountNotNumeric },
            new("throw syntax error", "{{#ds.Items}}x", null) { ExpectedError = TemplateErrorCode.TemplateSyntaxErrors },
        ];

        [TestCaseSource(nameof(Cases))]
        public void TextRenderer(Case testCase)
        {
            Check(testCase, RenderText);
        }

        [TestCaseSource(nameof(Cases))]
        public void WordRenderer(Case testCase)
        {
            Check(testCase, RenderDocx);
        }

        private static void Check(Case testCase, Func<Case, string> render)
        {
            if (testCase.ExpectedError != TemplateErrorCode.None)
            {
                var e = Assert.Catch<TemplateException>(() => render(testCase));
                Assert.That(e.ErrorCode, Is.EqualTo(testCase.ExpectedError));
            }
            else
            {
                Assert.That(Normalize(render(testCase)), Is.EqualTo(Normalize(testCase.Expected)));
            }
        }

        private static string Normalize(string text)
        {
            return text.Replace("\r\n", "\n").TrimEnd('\n');
        }

        private static string RenderText(Case testCase)
        {
            var template = new TextTemplate(testCase.Template, new TextTemplateSettings
            {
                Escaping = TextEscaping.None,
                Culture = testCase.Culture,
                UiCulture = CultureInfo.InvariantCulture,
                BindingErrorHandling = testCase.ErrorHandling,
                IgnoreLineBreaksAroundTags = testCase.IgnoreLineBreaksAroundTags
            });
            template.BindModel("ds", testCase.Model());
            return template.Process();
        }

        private static string RenderDocx(Case testCase)
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var body = new Body(testCase.Template.Split('\n').Select(line => new Paragraph(new Run(new WordText(line) { Space = SpaceProcessingModeValues.Preserve }))));
                wp.AddMainDocumentPart().Document = new Document(body);
                wp.Save();
            }
            stream.Position = 0;
            using var template = new DocxTemplate(stream, new ProcessSettings
            {
                Culture = testCase.Culture,
                UiCulture = CultureInfo.InvariantCulture,
                BindingErrorHandling = testCase.ErrorHandling,
                IgnoreLineBreaksAroundTags = testCase.IgnoreLineBreaksAroundTags
            });
            template.BindModel("ds", testCase.Model());
            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var paragraphs = document.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>()
                .Select(p => string.Concat(p.Descendants().Select(e => e switch
                {
                    WordText text => text.Text,
                    Break => "\n",
                    _ => string.Empty
                })));
            return string.Join("\n", paragraphs);
        }
    }
}
