using System.Collections.Concurrent;
using System.Globalization;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text.Test
{
    /// <summary>
    /// A mailing: a new template per recipient, all created with the same <see cref="TextTemplateSettings"/> - no
    /// leaks between recipients, parallel processing, collected errors and settings copied by each template.
    /// </summary>
    internal class MailingTest
    {
        private sealed class Company
        {
            public string Name { get; set; } = "Müller & \"Partner\" <AG>";
        }

        private sealed class Recipient
        {
            public int Id { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Email { get; set; }
            public List<string> Interests { get; set; } = [];
        }

        private sealed class Unsubscribe
        {
            public string Url { get; set; }
        }

        private sealed class Document
        {
            public DateTime Date { get; set; } = new(2026, 10, 20);
            public decimal Amount { get; set; } = 1234.5m;
        }

        private sealed class FailingFormatter : ITextFormatter
        {
            public bool CanHandle(Type type, string prefix)
            {
                return prefix.Equals("fail", StringComparison.OrdinalIgnoreCase);
            }

            public TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext)
            {
                throw new InvalidOperationException("formatter failed");
            }
        }

        private static readonly Recipient Person = new() { Id = 1, FirstName = "Ada", LastName = "Lovelace", Email = "ada@example.com", Interests = ["Math", "Engines"] };

        private static TextTemplateSettings Settings(TextEscaping escaping, BindingErrorHandling errorHandling = BindingErrorHandling.SkipBindingAndRemoveContent)
        {
            return new TextTemplateSettings
            {
                Escaping = escaping,
                Culture = CultureInfo.InvariantCulture,
                UiCulture = CultureInfo.InvariantCulture,
                BindingErrorHandling = errorHandling
            };
        }

        private static TextRenderResult Process(string template, TextTemplateSettings settings, string prefix, object model)
        {
            var textTemplate = new TextTemplate(template, settings);
            textTemplate.BindModel(prefix, model);
            return textTemplate.Process();
        }

        [Test]
        public void MailingSubjectAndBodyFromTheSameModels()
        {
            const string subjectText = "News von {{company.Name}} für {{(recipient.FirstName ?? recipient.LastName ?? recipient.Email)}}";
            const string bodyText =
                "<table><tr><td>" +
                "{?{recipient.FirstName != null}}Hallo {{recipient.FirstName}} {{recipient.LastName}}{{:}}Sehr geehrte {{(recipient.LastName ?? \"Damen und Herren\")}}{{/}}," +
                "</td></tr><tr><td>{{#recipient.Interests}}{{.}}{{:s:}}, {{/recipient.Interests}}</td></tr></table>" +
                "<p>{{company.Name}}</p><a href=\"{{unsubscribe.Url}}\">Abmelden</a>";
            var subjectSettings = Settings(TextEscaping.None);
            var bodySettings = Settings(TextEscaping.Html);
            var company = new Company();

            var expected = new[]
            {
                (Person, "News von Müller & \"Partner\" <AG> für Ada",
                    "<table><tr><td>Hallo Ada Lovelace,</td></tr><tr><td>Math, Engines</td></tr></table>"),
                (new Recipient { Id = 2, LastName = "Muster & Soehne AG", Email = "info@muster.ch" }, "News von Müller & \"Partner\" <AG> für Muster & Soehne AG",
                    "<table><tr><td>Sehr geehrte Muster &amp; Soehne AG,</td></tr><tr><td></td></tr></table>"),
                (new Recipient { Id = 3, Email = "lead@example.com" }, "News von Müller & \"Partner\" <AG> für lead@example.com",
                    "<table><tr><td>Sehr geehrte Damen und Herren,</td></tr><tr><td></td></tr></table>")
            };

            Assert.Multiple(() =>
            {
                foreach (var (recipient, expectedSubject, expectedTable) in expected)
                {
                    var subject = new TextTemplate(subjectText, subjectSettings);
                    var body = new TextTemplate(bodyText, bodySettings);
                    foreach (var template in new[] { subject, body })
                    {
                        template.BindModel("recipient", recipient);
                        template.BindModel("company", company);
                        template.BindModel("unsubscribe", new Unsubscribe { Url = $"https://example.com/unsubscribe?id={recipient.Id}&token=a\"b" });
                    }
                    var subjectResult = subject.Process();
                    var bodyResult = body.Process();
                    Assert.That(subjectResult.Text, Is.EqualTo(expectedSubject));
                    Assert.That(bodyResult.Text, Is.EqualTo(expectedTable +
                        "<p>Müller &amp; &quot;Partner&quot; &lt;AG&gt;</p>" +
                        $"<a href=\"https://example.com/unsubscribe?id={recipient.Id}&amp;token=a&quot;b\">Abmelden</a>"));
                    Assert.That(subjectResult.HasErrors || bodyResult.HasErrors, Is.False);
                }
            });
        }

        [TestCase(null, "Damen und Herren|Damen und Herren|ohne|ohne")]
        [TestCase("", "|Damen und Herren|ohne|ohne")]
        [TestCase("Ada", "Ada|Ada|mit|mit")]
        public void FallbackAndConditionOnMissingValues(string firstName, string expected)
        {
            const string template = "{{(recipient.FirstName ?? \"Damen und Herren\")}}" +
                                    "|{?{recipient.FirstName != null && recipient.FirstName != \"\"}}{{recipient.FirstName}}{{:}}Damen und Herren{{/}}" +
                                    "|{?{recipient.FirstName != null && recipient.FirstName != \"\"}}mit{{:}}ohne{{/}}" +
                                    "|{{#switch: recipient.FirstName}}{{#case: 'Ada'}}mit{{/}}{{#default}}ohne{{/}}{{/}}";
            var result = Process(template, Settings(TextEscaping.None), "recipient", new Recipient { FirstName = firstName });
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo(expected));
                Assert.That(result.Errors, Is.Empty);
            });
        }

        [Test]
        public void FallbackAndConditionOnMissingProperty()
        {
            // a dictionary without the key: the expression sees null, as with SkipBindingAndRemoveContent in a Word template
            const string template = "{{(recipient.FirstName ?? \"Damen und Herren\")}}|{?{recipient.FirstName != null}}mit{{:}}ohne{{/}}";
            var result = Process(template, Settings(TextEscaping.None), "recipient", new Dictionary<string, object> { ["Email"] = "lead@example.com" });
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo("Damen und Herren|ohne"));
                Assert.That(result.Errors, Is.Empty);
            });
        }

        [Test]
        public void NoValueLeaksIntoTheNextRecipient()
        {
            const string template = "[{{recipient.FirstName}}]" +
                                    "[{{(recipient.FirstName ?? \"-\")}}]" +
                                    "[{?{recipient.FirstName != null}}{{recipient.FirstName}}{{/}}]" +
                                    "[{{#recipient.Interests}}{{.}}:{{(..FirstName ?? \"-\")}}{?{..FirstName != null}}!{{/}}{{:s:}},{{/recipient.Interests}}]";
            var settings = Settings(TextEscaping.Html);
            var a = new Recipient { FirstName = "Ada", Interests = ["x", "y"] };
            string Render(Recipient recipient)
            {
                return Process(template, settings, "recipient", recipient).Text;
            }

            Assert.Multiple(() =>
            {
                Assert.That(Render(a), Is.EqualTo("[Ada][Ada][Ada][x:Ada!,y:Ada!]"));
                Assert.That(Render(new Recipient { Interests = ["z"] }), Is.EqualTo("[][-][][z:-]"));
                Assert.That(Render(new Recipient()), Is.EqualTo("[][-][][]"));
                Assert.That(Render(a), Is.EqualTo("[Ada][Ada][Ada][x:Ada!,y:Ada!]"));
            });
        }

        [Test]
        public void ValuesOfAnotherTypeAreEvaluatedWithTheirOwnExpression()
        {
            // the cached expressions of the settings are parsed per type of the values
            const string template = "{{(recipient.Value + 1)}}|{?{recipient.Value != null}}set{{:}}null{{/}}";
            var settings = Settings(TextEscaping.None);
            object[] values = [1, 2.5, 3m, "x", null, 4L];
            var results = new string[values.Length * 20];
            Parallel.For(0, results.Length, i => results[i] = Process(template, settings, "recipient", new Dictionary<string, object> { ["Value"] = values[i % values.Length] }).Text);
            string[] expected = ["2|set", "3.5|set", "4|set", "x1|set", "|null", "5|set"];
            Assert.That(results, Is.EqualTo(Enumerable.Range(0, results.Length).Select(i => expected[i % values.Length])));
        }

        [Test]
        public void TemplatesProcessedInParallelMatchSequentialProcessing()
        {
            const string template = "<p>{?{recipient.FirstName != null}}Hallo {{recipient.FirstName}}{{:}}Guten Tag{{/}} #{{recipient.Id}}</p>" +
                                    "<p>{{(recipient.FirstName ?? \"Damen und Herren\")}}</p>" +
                                    "<ul>{{#recipient.Interests}}<li>{{recipient.Interests._Idx}}/{{.}}/{{..Id}}</li>{{/recipient.Interests}}</ul>" +
                                    "{{#switch: recipient.FirstName}}{{#case: 'F4'}}four{{/}}{{#default}}other{{/}}{{/}}";
            var recipients = Enumerable.Range(0, 3000)
                .Select(i => new Recipient
                {
                    Id = i,
                    FirstName = i % 3 == 0 ? null : $"F{i}",
                    Interests = Enumerable.Range(0, i % 4).Select(x => $"I{i}-{x}").ToList()
                })
                .ToList();

            var sequentialSettings = Settings(TextEscaping.Html);
            var sequential = recipients.Select(r => Process(template, sequentialSettings, "recipient", r).Text).ToList();
            var parallelSettings = Settings(TextEscaping.Html);
            var parallel = new string[recipients.Count];
            var errors = new ConcurrentBag<TextRenderError>();
            Parallel.For(0, recipients.Count, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i =>
            {
                var result = Process(template, parallelSettings, "recipient", recipients[i]);
                parallel[i] = result.Text;
                foreach (var error in result.Errors)
                {
                    errors.Add(error);
                }
            });

            Assert.Multiple(() =>
            {
                Assert.That(parallel, Is.EqualTo(sequential));
                Assert.That(errors, Is.Empty);
                Assert.That(sequential[4], Is.EqualTo("<p>Hallo F4 #4</p><p>F4</p><ul></ul>four"));
                Assert.That(sequential[6], Is.EqualTo("<p>Guten Tag #6</p><p>Damen und Herren</p><ul><li>1/I6-0/6</li><li>2/I6-1/6</li></ul>other"));
            });
        }

        [Test]
        public void CollectedErrorsAreReportedWithPositionAndLeftOutOfTheText()
        {
            const string template = "Hallo {{recipient.Nickname}}! {{recipient.FirstName}:shout()} {{(recipient.FirstName / 2)}} {?{recipient.FirstName > }}x{{/}}Ende";
            var result = Process(template, Settings(TextEscaping.Html), "recipient", Person);
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo("Hallo !   Ende"));
                Assert.That(result.Errors.Select(x => x.Tag), Is.EqualTo(
                [
                    "{{recipient.Nickname}}", "{{recipient.FirstName}:shout()}", "{{(recipient.FirstName / 2)}}", "{?{recipient.FirstName > }}"
                ]));
                Assert.That(result.Errors.Select(x => x.Index), Is.EqualTo(result.Errors.Select(x => template.IndexOf(x.Tag, StringComparison.Ordinal))));
                Assert.That(result.Errors.Select(x => x.Length), Is.EqualTo(result.Errors.Select(x => x.Tag.Length)));
                Assert.That(result.Errors[0].Exception.ErrorCode, Is.EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
                Assert.That(result.Errors[1].Exception.ErrorCode, Is.EqualTo(TemplateErrorCode.UnknownFormatter));
                Assert.That(result.Errors[1].Exception.Message, Is.EqualTo("'{{recipient.FirstName}:shout()}': there is no formatter 'shout' for a value of type 'System.String'"));
                Assert.That(result.Errors[3].Exception.ErrorCode, Is.EqualTo(TemplateErrorCode.ErrorInCondition));
                Assert.That(result.Errors.All(x => !string.IsNullOrEmpty(x.Exception.Message)));
            });
        }

        [Test]
        public void ExceptionsOfExpressionsAndFormattersAreCollected()
        {
            const string template = "A{{(recipient.Interests[5])}}B{?{10 / recipient.Id > 1}}x{{/}}C{{recipient.FirstName}:fail()}D";
            var settings = Settings(TextEscaping.None);
            settings.RegisterFormatter(new FailingFormatter());
            var result = Process(template, settings, "recipient", new Recipient { Id = 0, FirstName = "Ada" });
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo("ABCD"));
                Assert.That(result.Errors.Select(x => x.Tag), Is.EqualTo(["{{(recipient.Interests[5])}}", "{?{10 / recipient.Id > 1}}", "{{recipient.FirstName}:fail()}"]));
                Assert.That(result.Errors[1].Exception.InnerException, Is.InstanceOf<DivideByZeroException>());
                Assert.That(result.Errors[2].Exception.InnerException, Is.InstanceOf<InvalidOperationException>().With.Message.EqualTo("formatter failed"));
            });
        }

        [TestCase(BindingErrorHandling.ThrowException)]
        [TestCase(BindingErrorHandling.SkipBindingAndRemoveContent)]
        public void TextThatLooksLikeATagIsReportedByTheValidationOnly(BindingErrorHandling errorHandling)
        {
            var template = new TextTemplate("Hello {{recipient.<b>First</b>Name}} {{recipient.LastName}}", Settings(TextEscaping.None, errorHandling));
            template.BindModel("recipient", Person);
            var result = template.Process();
            var warning = template.ValidateTemplateSyntax().Single();
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo("Hello {{recipient.<b>First</b>Name}} Lovelace"));
                // an error of the template, not of the recipient
                Assert.That(result.HasErrors, Is.False);
                Assert.That((warning.Severity, warning.ErrorCode), Is.EqualTo((TemplateSyntaxErrorSeverity.Warning, TemplateErrorCode.InvalidTag)));
                Assert.That(warning.Tag, Is.EqualTo("{{recipient.<b>First</b>Name}}"));
                Assert.That(warning.Index, Is.EqualTo(6));
            });
        }

        [Test]
        public void ClosingBracesWithoutTagAreText()
        {
            // the end of a CSS media query: "}" and "}" on the next line
            const string css = "<style>@media (max-width: 620px) {\n  .btn {\n    padding: 0;\n  }\n}\n</style>";
            var template = new TextTemplate(css + "Hello {{recipient.LastName}} }}", Settings(TextEscaping.Html, BindingErrorHandling.ThrowException));
            template.BindModel("recipient", Person);
            var result = template.Process();
            var warnings = template.ValidateTemplateSyntax();
            Assert.Multiple(() =>
            {
                Assert.That(result.Text, Is.EqualTo(css + "Hello Lovelace }}"));
                Assert.That(result.HasErrors, Is.False);
                Assert.That(warnings.Select(x => (x.Severity, x.ErrorCode)), Is.All.EqualTo((TemplateSyntaxErrorSeverity.Warning, TemplateErrorCode.UnexpectedClosingBraces)));
                Assert.That(warnings, Has.Count.EqualTo(2));
            });
        }

        [Test]
        public void CulturePerSettings()
        {
            const string template = "{{document.Date}:f(d)}|{{document.Amount}:f(N2)}|{{document.Amount}}";
            string Render(string culture)
            {
                var settings = Settings(TextEscaping.None);
                settings.Culture = new CultureInfo(culture);
                return Process(template, settings, "document", new Document()).Text;
            }

            string Expected(string culture)
            {
                var c = new CultureInfo(culture);
                return $"{new DateTime(2026, 10, 20).ToString("d", c)}|{1234.5m.ToString("N2", c)}|{1234.5m.ToString(c)}";
            }

            Assert.Multiple(() =>
            {
                Assert.That(Render("de-CH"), Is.EqualTo(Expected("de-CH")).And.StartsWith("20.10.2026|1").And.EndsWith("234.50|1234.5"));
                Assert.That(Render("fr-CH"), Is.EqualTo(Expected("fr-CH")).And.StartsWith("20.10.2026|1").And.EndsWith("234,50|1234,5"));
                Assert.That(Render("en-US"), Is.EqualTo("10/20/2026|1,234.50|1234.5"));
            });
        }

        [Test]
        public void SettingsChangedAfterCreatingATemplateDoNotAffectIt()
        {
            const string template = "{{recipient.FirstName}:fail()}|{{(recipient.Id + 1)}}";
            var settings = Settings(TextEscaping.None);
            var before = new TextTemplate(template, settings);
            settings.RegisterFormatter(new FailingFormatter());
            settings.Culture = new CultureInfo("de-CH");
            var after = new TextTemplate(template, settings);
            before.BindModel("recipient", Person);
            after.BindModel("recipient", Person);
            var beforeResult = before.Process();
            var afterResult = after.Process();
            Assert.Multiple(() =>
            {
                Assert.That(beforeResult.Errors.Single().Exception.ErrorCode, Is.EqualTo(TemplateErrorCode.UnknownFormatter));
                Assert.That(afterResult.Errors.Single().Exception.InnerException.Message, Is.EqualTo("formatter failed"));
                // the expression parsed for the first template is reused by the second
                Assert.That(beforeResult.Text, Is.EqualTo("|2"));
                Assert.That(afterResult.Text, Is.EqualTo("|2"));
            });
        }
    }
}
