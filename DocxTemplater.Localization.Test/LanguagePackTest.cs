using System.Collections;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace DocxTemplater.Localization.Test
{
    /// <summary>
    /// The same checks for every language package: complete, consistent with the English placeholders and
    /// applied end-to-end through <see cref="ProcessSettings.UiCulture"/>.
    /// </summary>
    [TestFixture]
    public class LanguagePackTest
    {
        public sealed record LanguagePack(string Name, CultureInfo Culture, CultureInfo SpecificCulture,
            IReadOnlyDictionary<TemplateErrorCode, string> Formats, Func<TemplateErrorMessages, TemplateErrorMessages> Add)
        {
            public override string ToString()
            {
                return Name;
            }
        }

        public static IEnumerable Packs
        {
            get
            {
                yield return new TestCaseData(new LanguagePack("German", GermanErrorMessages.Culture, new CultureInfo("de-AT"), GermanErrorMessages.Formats, m => m.AddGerman()));
                yield return new TestCaseData(new LanguagePack("SwissGerman", SwissGermanErrorMessages.Culture, new CultureInfo("de-CH"), SwissGermanErrorMessages.Formats, m => m.AddSwissGerman()));
                yield return new TestCaseData(new LanguagePack("French", FrenchErrorMessages.Culture, new CultureInfo("fr-CH"), FrenchErrorMessages.Formats, m => m.AddFrench()));
                yield return new TestCaseData(new LanguagePack("Italian", ItalianErrorMessages.Culture, new CultureInfo("it-CH"), ItalianErrorMessages.Formats, m => m.AddItalian()));
            }
        }

        private static readonly Regex PlaceholderRegex = new(@"\{(\d+)\}", RegexOptions.Compiled);

        private static IEnumerable<string> Placeholders(string format)
        {
            // literal braces are doubled - remove them before looking for the placeholders
            return PlaceholderRegex.Matches(format.Replace("{{", "").Replace("}}", "")).Select(m => m.Value).Distinct().OrderBy(x => x);
        }

        private static DocxTemplate BuildTemplate(ProcessSettings settings, params string[] paragraphs)
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var body = new Body(paragraphs.Select(p => new Paragraph(new Run(new Text(p)))));
                wp.AddMainDocumentPart().Document = new Document(body);
                wp.Save();
            }
            stream.Position = 0;
            return new DocxTemplate(stream, settings);
        }

        [TestCaseSource(nameof(Packs))]
        public void AllCodesAreTranslated(LanguagePack pack)
        {
            var codes = Enum.GetValues<TemplateErrorCode>().Where(x => x != TemplateErrorCode.None);

            Assert.Multiple(() =>
            {
                foreach (var code in codes)
                {
                    Assert.That(pack.Formats.ContainsKey(code), $"{code} is not translated");
                }
                Assert.That(pack.Formats.Keys, Is.EquivalentTo(TemplateErrorMessages.English.Keys));
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void PlaceholdersMatchEnglish(LanguagePack pack)
        {
            Assert.Multiple(() =>
            {
                foreach (var (code, format) in pack.Formats)
                {
                    Assert.That(Placeholders(format), Is.EqualTo(Placeholders(TemplateErrorMessages.English[code])), $"{code}: placeholders differ from English");
                }
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void AllFormatsCanBeFormatted(LanguagePack pack)
        {
            var messages = pack.Add(new TemplateErrorMessages());
            object[] args = ["a", "b", "c"];

            Assert.Multiple(() =>
            {
                foreach (var (code, format) in pack.Formats)
                {
                    // string.Format throws on a malformed format - and the provider would silently fall back to English
                    var expected = string.Format(pack.Culture, format.Replace("\n", Environment.NewLine), args);
                    Assert.That(messages.Format(code, pack.SpecificCulture, args), Is.EqualTo(expected), $"{code} fell back to English");
                }
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void BindingError_IsTranslated(LanguagePack pack)
        {
            var settings = new ProcessSettings
            {
                UiCulture = pack.SpecificCulture,
                ErrorMessages = pack.Add(new TemplateErrorMessages()),
                BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument
            };
            using var template = BuildTemplate(settings, "{{Foo}}");

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            var expected = string.Format(pack.Culture, pack.Formats[TemplateErrorCode.ModelNotFound], "Foo");
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Does.Contain(expected));
        }

        [TestCaseSource(nameof(Packs))]
        public void SyntaxError_IsTranslated(LanguagePack pack)
        {
            var settings = new ProcessSettings
            {
                UiCulture = pack.SpecificCulture,
                ErrorMessages = pack.Add(new TemplateErrorMessages())
            };
            using var template = BuildTemplate(settings, "{{#Items}}", "{{.}}");

            var errors = template.ValidateTemplateSyntax();
            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            var expectedLine = string.Format(pack.Culture, pack.Formats[TemplateErrorCode.BlockNotClosed], "{{#Items}}");
            Assert.Multiple(() =>
            {
                Assert.That(errors, Has.Count.EqualTo(1));
                Assert.That(errors[0].ErrorCode, Is.EqualTo(TemplateErrorCode.BlockNotClosed));
                Assert.That(errors[0].Message, Is.EqualTo(expectedLine));
                Assert.That(errors[0].GetMessage(CultureInfo.InvariantCulture), Is.EqualTo("'{{#Items}}' is not closed"));
                Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
                Assert.That(ex.Message, Does.Contain(expectedLine));
                Assert.That(ex.Message, Does.Not.Contain("Template syntax errors"));
                Assert.That(ex.GetMessage(CultureInfo.InvariantCulture), Does.Contain("Template syntax errors").And.Contain("'{{#Items}}' is not closed"));
            });
        }

        [Test]
        public void SwissGerman_UsesSsInsteadOfSharpS()
        {
            Assert.Multiple(() =>
            {
                Assert.That(GermanErrorMessages.Formats[TemplateErrorCode.ModelPrefixAlreadyBound], Does.Contain("Groß-"));
                Assert.That(SwissGermanErrorMessages.Formats[TemplateErrorCode.ModelPrefixAlreadyBound], Does.Contain("Gross-"));
                Assert.That(SwissGermanErrorMessages.Formats.Values, Has.None.Contains("ß"));
                Assert.That(GermanErrorMessages.Formats.Values, Has.Some.Contains("ß"), "the German texts should use standard spelling");
            });
        }

        [Test]
        public void SpecificCulture_FallsBackToNeutralCulture()
        {
            var germanOnly = new TemplateErrorMessages().AddGerman();
            var both = new TemplateErrorMessages().AddGerman().AddSwissGerman();
            var swissOnly = new TemplateErrorMessages().AddSwissGerman();

            Assert.Multiple(() =>
            {
                // de-CH -> de
                Assert.That(germanOnly.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-CH"), "x"), Does.Contain("Groß-"));
                // de-CH texts win over the de fallback, other German cultures still use de
                Assert.That(both.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-CH"), "x"), Does.Contain("Gross-"));
                Assert.That(both.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-DE"), "x"), Does.Contain("Groß-"));
                Assert.That(both.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-AT"), "x"), Does.Contain("Groß-"));
                Assert.That(both.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de"), "x"), Does.Contain("Groß-"));
                // de-CH texts are not used for other German cultures - they fall back to English
                Assert.That(swissOnly.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("de-DE"), "X"), Is.EqualTo("Model X not found"));
                Assert.That(swissOnly.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("de-CH"), "X"), Is.EqualTo("Modell X nicht gefunden"));
                // fr-CH -> fr
                Assert.That(new TemplateErrorMessages().AddFrench().Format(TemplateErrorCode.ModelNotFound, new CultureInfo("fr-CH"), "X"), Is.EqualTo("Modèle X introuvable"));
            });
        }

        [Test]
        public void Default_CanRegisterAllLanguages()
        {
            var messages = new TemplateErrorMessages().AddGerman().AddSwissGerman().AddFrench().AddItalian();

            Assert.Multiple(() =>
            {
                string[] languages = ["de", "de-CH", "fr", "it"];
                Assert.That(messages.Languages, Is.EquivalentTo(languages));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("de-AT"), "X"), Is.EqualTo("Modell X nicht gefunden"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("fr-CA"), "X"), Is.EqualTo("Modèle X introuvable"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("it-IT"), "X"), Is.EqualTo("Modello X non trovato"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("es-ES"), "X"), Is.EqualTo("Model X not found"));
            });
        }
    }
}
