using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class ErrorMessagesTest
    {
        private static readonly CultureInfo English = new("en-US");
        private static readonly CultureInfo French = new("fr");
        private static readonly CultureInfo SwissFrench = new("fr-CH");

        // a partial language: codes that are missing fall back to English
        private static TemplateErrorMessages CreateFrenchMessages()
        {
            return new TemplateErrorMessages().AddLanguage(French, new Dictionary<TemplateErrorCode, string>
            {
                [TemplateErrorCode.ModelNotFound] = "Modèle {0} introuvable",
                [TemplateErrorCode.ErrorInCondition] = "{0} dans la condition '{1}'",
                [TemplateErrorCode.NoMatchingOpeningTag] = "'{0}' n'a pas de balise ouvrante",
                [TemplateErrorCode.SyntaxErrorLocation] = "{0} : {1} (près de '{2}')"
            });
        }

        private static ProcessSettings FrenchSettings(BindingErrorHandling errorHandling)
        {
            return new ProcessSettings
            {
                UiCulture = SwissFrench,
                ErrorMessages = CreateFrenchMessages(),
                BindingErrorHandling = errorHandling
            };
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

        [Test]
        public void BindingError_CarriesErrorCodeAndArguments()
        {
            using var template = BuildTemplate(new ProcessSettings { UiCulture = English }, "{{ds.Foo}}");
            template.BindModel("ds", new { Bar = 1 });

            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            Assert.Multiple(() =>
            {
                Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.PlaceholderNotReplaced));
                Assert.That(ex.Arguments[0], Is.EqualTo("{{ds.Foo}}"));
                Assert.That(ex.Message, Does.StartWith("'{{ds.Foo}}' could not be replaced"));
                var inner = ex.InnerException as OpenXmlTemplateException;
                Assert.That(inner, Is.Not.Null);
                Assert.That(inner.ErrorCode, Is.EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
                Assert.That(inner.Arguments[0], Is.EqualTo("Foo"));
                Assert.That(inner.Message, Does.StartWith("Property 'Foo' not found in 'ds.Foo'"));
            });
        }

        [Test]
        public void UiCulture_SelectsLanguageOfExceptionMessage()
        {
            using var template = BuildTemplate(FrenchSettings(BindingErrorHandling.ThrowException), "{{Foo}}");

            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            var inner = (OpenXmlTemplateException)ex.InnerException;
            Assert.Multiple(() =>
            {
                // fr-CH falls back to fr
                Assert.That(inner.ErrorCode, Is.EqualTo(TemplateErrorCode.ModelNotFound));
                Assert.That(inner.Message, Is.EqualTo("Modèle Foo introuvable"));
                Assert.That(inner.GetMessage(English), Is.EqualTo("Model Foo not found"));
                // GetMessage(culture) uses the TemplateErrorMessages instance the exception was created with, not Default
                Assert.That(inner.GetMessage(French), Is.EqualTo("Modèle Foo introuvable"));
                // not translated - falls back to English
                Assert.That(ex.Message, Does.Contain("could not be replaced"));
            });
        }

        [Test]
        public void InvariantCultureTexts_OverrideBuiltInEnglish()
        {
            var messages = new TemplateErrorMessages()
                .AddLanguage(CultureInfo.InvariantCulture, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.ModelNotFound] = "Model '{0}' is missing" });

            Assert.Multiple(() =>
            {
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, CultureInfo.InvariantCulture, "X"), Is.EqualTo("Model 'X' is missing"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, English, "X"), Is.EqualTo("Model 'X' is missing"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, SwissFrench, "X"), Is.EqualTo("Model 'X' is missing"), "no French texts - the override is the last fallback before English");
                Assert.That(messages.Format(TemplateErrorCode.BlockNotClosed, English, "x"), Is.EqualTo("'x' is not closed"));
            });
        }

        [Test]
        public void SyntaxError_WithoutCode_KeepsFreeTextMessage()
        {
            var error = new TemplateSyntaxError(TemplateSyntaxErrorSeverity.Error, "Body", "{{x}}", "ctx",
                TemplateErrorCode.None, null, null, null, message: "free text");

            Assert.Multiple(() =>
            {
                Assert.That(error.Message, Is.EqualTo("free text"));
                Assert.That(error.GetMessage(French), Is.EqualTo("free text"));
                Assert.That(error.ToString(), Is.EqualTo("Body: free text (near 'ctx')"));
            });
        }

        [Test]
        public void HighlightErrorsInDocument_WritesTranslatedMessage()
        {
            using var template = BuildTemplate(FrenchSettings(BindingErrorHandling.HighlightErrorsInDocument), "{{Foo}}");

            var result = template.Process();

            using var document = WordprocessingDocument.Open(result, false);
            Assert.That(document.MainDocumentPart.Document.Body.InnerText, Does.Contain("Modèle Foo introuvable"));
        }

        [Test]
        public void SyntaxError_CarriesErrorCodeAndTranslates()
        {
            using var template = BuildTemplate(FrenchSettings(BindingErrorHandling.ThrowException), "{{/}}");

            var errors = template.ValidateTemplateSyntax();

            Assert.That(errors, Has.Count.EqualTo(1));
            var error = errors[0];
            Assert.Multiple(() =>
            {
                Assert.That(error.ErrorCode, Is.EqualTo(TemplateErrorCode.NoMatchingOpeningTag));
                Assert.That(error.Arguments[0], Is.EqualTo("{{/}}"));
                Assert.That(error.Message, Is.EqualTo("'{{/}}' n'a pas de balise ouvrante"));
                Assert.That(error.GetMessage(English), Is.EqualTo("'{{/}}' has no matching opening tag"));
                Assert.That(error.ToString(), Is.EqualTo("Body : '{{/}}' n'a pas de balise ouvrante (près de '{{/}}')"));
                Assert.That(error.ToString(English), Is.EqualTo("Body: '{{/}}' has no matching opening tag (near '{{/}}')"));
            });
        }

        [Test]
        public void Process_WithSyntaxErrors_ExceptionTranslatesNestedErrors()
        {
            using var template = BuildTemplate(FrenchSettings(BindingErrorHandling.ThrowException), "{{/}}");

            var ex = Assert.Throws<OpenXmlTemplateException>(() => template.Process());

            Assert.Multiple(() =>
            {
                Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.TemplateSyntaxErrors));
                // the header is not translated, the lines are
                Assert.That(ex.Message, Is.EqualTo($"Template syntax errors:{Environment.NewLine}Body : '{{{{/}}}}' n'a pas de balise ouvrante (près de '{{{{/}}}}')"));
                Assert.That(ex.GetMessage(English), Is.EqualTo($"Template syntax errors:{Environment.NewLine}Body: '{{{{/}}}}' has no matching opening tag (near '{{{{/}}}}')"));
            });
        }

        [Test]
        public void NestedError_IsTranslatedAsWhole()
        {
            var settings = FrenchSettings(BindingErrorHandling.ThrowException);
            var inner = OpenXmlTemplateException.Create(settings, TemplateErrorCode.ModelNotFound, "Foo");

            var outer = OpenXmlTemplateException.Create(settings, TemplateErrorCode.ErrorInCondition, inner, "Foo > 1");

            Assert.Multiple(() =>
            {
                Assert.That(outer.Message, Is.EqualTo("Modèle Foo introuvable dans la condition 'Foo > 1'"));
                Assert.That(outer.GetMessage(English), Is.EqualTo("Model Foo not found in condition 'Foo > 1'"));
                Assert.That(outer.InnerException, Is.Null, "the inner error is an argument, not the cause");
            });
        }

        [Test]
        public void MissingTranslation_FallsBackToParentCultureAndEnglish()
        {
            var messages = CreateFrenchMessages()
                .AddLanguage(SwissFrench, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.BlockNotClosed] = "'{0}' n'est pas fermé (CH)" });

            Assert.Multiple(() =>
            {
                Assert.That(messages.Format(TemplateErrorCode.BlockNotClosed, SwissFrench, "{{#x}}"), Is.EqualTo("'{{#x}}' n'est pas fermé (CH)"));
                Assert.That(messages.Format(TemplateErrorCode.BlockNotClosed, French, "{{#x}}"), Is.EqualTo("'{{#x}}' is not closed"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, SwissFrench, "X"), Is.EqualTo("Modèle X introuvable"));
                Assert.That(messages.Format(TemplateErrorCode.InvalidTag, SwissFrench, "{{x"), Is.EqualTo("Invalid tag '{{x'"));
                Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, CultureInfo.InvariantCulture, "X"), Is.EqualTo("Model X not found"));
                string[] languages = ["fr", "fr-CH"];
                Assert.That(messages.Languages, Is.EquivalentTo(languages));
            });
        }

        [Test]
        public void AddLanguage_MergesDictionaries()
        {
            var messages = new TemplateErrorMessages()
                .AddLanguage(French, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.ModelNotFound] = "A {0}" })
                .AddLanguage(French, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.BlockNotClosed] = "B {0}" })
                .AddLanguage(French, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.ModelNotFound] = "C {0}" });

            Assert.Multiple(() =>
            {
                Assert.That(messages.GetFormat(TemplateErrorCode.ModelNotFound, French), Is.EqualTo("C {0}"));
                Assert.That(messages.GetFormat(TemplateErrorCode.BlockNotClosed, French), Is.EqualTo("B {0}"));
                string[] languages = ["fr"];
                Assert.That(messages.Languages, Is.EqualTo(languages));
            });
        }

        [Test]
        public void InvalidCustomFormat_FallsBackToEnglish()
        {
            var messages = new TemplateErrorMessages()
                .AddLanguage(French, new Dictionary<TemplateErrorCode, string> { [TemplateErrorCode.ModelNotFound] = "Modèle {0} {1}" });

            Assert.That(messages.Format(TemplateErrorCode.ModelNotFound, French, "X"), Is.EqualTo("Model X not found"));
        }

        [Test]
        public void EnglishTexts_ExistForAllCodes()
        {
            var codes = Enum.GetValues<TemplateErrorCode>().Where(x => x != TemplateErrorCode.None).ToList();
            Assert.Multiple(() =>
            {
                foreach (var code in codes)
                {
                    Assert.That(TemplateErrorMessages.English.ContainsKey(code), $"{code} has no English text");
                    // a malformed English format makes Format fall back to "<code>: a, b, c"
                    var message = TemplateErrorMessages.Default.Format(code, CultureInfo.InvariantCulture, "a", "b", "c");
                    Assert.That(message, Does.Not.StartWith(code.ToString()), $"{code} could not be formatted");
                }
            });
            Assert.That(TemplateErrorMessages.English, Has.Count.EqualTo(codes.Count));
        }

        [Test]
        public void FreeTextException_HasNoCode()
        {
            var ex = new OpenXmlTemplateException("free text");

            Assert.Multiple(() =>
            {
                Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.None));
                Assert.That(ex.Arguments, Is.Empty);
                Assert.That(ex.GetMessage(French), Is.EqualTo("free text"));
            });
        }

        [Test]
        public void ProcessSettings_Defaults()
        {
            var settings = new ProcessSettings();

            Assert.Multiple(() =>
            {
                Assert.That(settings.UiCulture, Is.EqualTo(CultureInfo.CurrentUICulture));
                Assert.That(settings.ErrorMessages, Is.SameAs(TemplateErrorMessages.Default));
            });
        }

        [Test]
        public void Create_WithoutSettings_IsEnglish()
        {
            var ex = OpenXmlTemplateException.Create(null, TemplateErrorCode.ModelNotFound, "X");

            Assert.That(ex.Message, Is.EqualTo("Model X not found"));
        }
    }
}
