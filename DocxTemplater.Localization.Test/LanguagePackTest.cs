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
    /// The same checks for every language package: complete, consistent with the English placeholders, loaded
    /// automatically by convention and applied end-to-end through <see cref="ProcessSettings.UiCulture"/>.
    /// </summary>
    [TestFixture]
    public class LanguagePackTest
    {
        public sealed record LanguagePack(ITemplateLanguagePack Pack, CultureInfo SpecificCulture)
        {
            public override string ToString()
            {
                return Pack.Culture.Name;
            }
        }

        public static IEnumerable Packs
        {
            get
            {
                yield return new TestCaseData(new LanguagePack(new GermanErrorMessages(), new CultureInfo("de-AT")));
                yield return new TestCaseData(new LanguagePack(new SwissGermanErrorMessages(), new CultureInfo("de-CH")));
                yield return new TestCaseData(new LanguagePack(new FrenchErrorMessages(), new CultureInfo("fr-CH")));
                yield return new TestCaseData(new LanguagePack(new ItalianErrorMessages(), new CultureInfo("it-CH")));
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

        private static string Expected(LanguagePack pack, TemplateErrorCode code, params object[] args)
        {
            return string.Format(pack.Pack.Culture, pack.Pack.Formats[code].Replace("\n", Environment.NewLine), args);
        }

        [TestCaseSource(nameof(Packs))]
        public void AllCodesAreTranslated(LanguagePack pack)
        {
            var codes = Enum.GetValues<TemplateErrorCode>().Where(x => x != TemplateErrorCode.None);

            Assert.Multiple(() =>
            {
                foreach (var code in codes)
                {
                    Assert.That(pack.Pack.Formats.ContainsKey(code), $"{code} is not translated");
                }
                Assert.That(pack.Pack.Formats.Keys, Is.EquivalentTo(TemplateErrorMessages.English.Keys));
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void PlaceholdersMatchEnglish(LanguagePack pack)
        {
            Assert.Multiple(() =>
            {
                foreach (var (code, format) in pack.Pack.Formats)
                {
                    Assert.That(Placeholders(format), Is.EqualTo(Placeholders(TemplateErrorMessages.English[code])), $"{code}: placeholders differ from English");
                }
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void AllFormatsCanBeFormatted(LanguagePack pack)
        {
            var messages = new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(pack.Pack);
            object[] args = ["a", "b", "c"];

            Assert.Multiple(() =>
            {
                foreach (var code in pack.Pack.Formats.Keys)
                {
                    // string.Format throws on a malformed format - and the provider would silently fall back to English
                    Assert.That(messages.Format(code, pack.SpecificCulture, args), Is.EqualTo(Expected(pack, code, args)), $"{code} fell back to English");
                }
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void Package_IsLoadedAutomaticallyByCultureName(LanguagePack pack)
        {
            // nothing registered - the package is found by its assembly name DocxTemplater.Localization.<culture>
            var messages = new TemplateErrorMessages();

            var message = messages.Format(TemplateErrorCode.ModelNotFound, pack.SpecificCulture, "X");

            Assert.Multiple(() =>
            {
                Assert.That(message, Is.EqualTo(Expected(pack, TemplateErrorCode.ModelNotFound, "X")));
                Assert.That(messages.Languages, Does.Contain(pack.Pack.Culture.Name));
            });
        }

        [TestCaseSource(nameof(Packs))]
        public void BindingError_IsTranslated(LanguagePack pack)
        {
            var settings = new ProcessSettings
            {
                UiCulture = pack.SpecificCulture,
                ErrorMessages = new TemplateErrorMessages(),
                BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument
            };
            using var template = BuildTemplate(settings, "{{Foo}}");

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Does.Contain(Expected(pack, TemplateErrorCode.ModelNotFound, "Foo")));
        }

        [TestCaseSource(nameof(Packs))]
        public void SyntaxError_IsTranslated(LanguagePack pack)
        {
            var settings = new ProcessSettings
            {
                UiCulture = pack.SpecificCulture,
                ErrorMessages = new TemplateErrorMessages()
            };
            using var template = BuildTemplate(settings, "{{#Items}}", "{{.}}");

            var errors = template.ValidateTemplateSyntax();
            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            var expectedLine = Expected(pack, TemplateErrorCode.BlockNotClosed, "{{#Items}}");
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
        public void AutoLoad_CanBeDisabled()
        {
            var messages = new TemplateErrorMessages { AutoLoadLanguagePackages = false };

            Assert.Multiple(() =>
            {
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("de-CH"), "X"), Is.EqualTo("Model X not found"));
                Assert.That(messages.Languages, Is.Empty);
                // explicit registration still works
                messages.AddLanguage(new SwissGermanErrorMessages());
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("de-CH"), "X"), Is.EqualTo("Modell X nicht gefunden"));
            });
        }

        [Test]
        public void AutoLoad_UnknownCulture_FallsBackToEnglishWithoutError()
        {
            var messages = new TemplateErrorMessages();

            Assert.Multiple(() =>
            {
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("es-ES"), "X"), Is.EqualTo("Model X not found"));
                // probed once, then cached - a second call must not try again (and must still be English)
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("es-ES"), "X"), Is.EqualTo("Model X not found"));
                Assert.That(messages.Languages, Is.Empty);
            });
        }

        [Test]
        public void SwissGerman_UsesSsInsteadOfSharpS()
        {
            var german = new GermanErrorMessages().Formats;
            var swiss = new SwissGermanErrorMessages().Formats;

            Assert.Multiple(() =>
            {
                Assert.That(german[TemplateErrorCode.ModelPrefixAlreadyBound], Does.Contain("Groß-"));
                Assert.That(swiss[TemplateErrorCode.ModelPrefixAlreadyBound], Does.Contain("Gross-"));
                Assert.That(swiss.Values, Has.None.Contains("ß"));
                Assert.That(german.Values, Has.Some.Contains("ß"), "the German texts should use standard spelling");
            });
        }

        [Test]
        public void SpecificCulture_FallsBackToNeutralCulture()
        {
            var germanOnly = new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(new GermanErrorMessages());
            var both = new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(new GermanErrorMessages()).AddLanguage(new SwissGermanErrorMessages());
            var swissOnly = new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(new SwissGermanErrorMessages());
            var frenchOnly = new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(new FrenchErrorMessages());

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
                Assert.That(frenchOnly.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("fr-CH"), "X"), Is.EqualTo("Modèle X introuvable"));
            });
        }

        [Test]
        public void AutoLoad_SpecificAndNeutralPackage()
        {
            // de-CH and de are both installed: de-CH gets the Swiss texts, de-AT the standard ones
            var messages = new TemplateErrorMessages();

            Assert.Multiple(() =>
            {
                Assert.That(messages.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-CH"), "x"), Does.Contain("Gross-"));
                Assert.That(messages.Format(TemplateErrorCode.ModelPrefixAlreadyBound, new CultureInfo("de-AT"), "x"), Does.Contain("Groß-"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("fr-CA"), "X"), Is.EqualTo("Modèle X introuvable"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, new CultureInfo("it-IT"), "X"), Is.EqualTo("Modello X non trovato"));
                string[] languages = ["de", "de-CH", "fr", "it"];
                Assert.That(messages.Languages, Is.EquivalentTo(languages));
            });
        }
    }
}
