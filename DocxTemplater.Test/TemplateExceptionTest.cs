using System.Globalization;

namespace DocxTemplater.Test
{
    internal class TemplateExceptionTest
    {
        private static readonly CultureInfo German = new("de");

        private static TemplateErrorMessages CreateGermanMessages()
        {
            return new TemplateErrorMessages { AutoLoadLanguagePackages = false }.AddLanguage(German, new Dictionary<TemplateErrorCode, string>
            {
                [TemplateErrorCode.ModelNotFound] = "Modell {0} nicht gefunden"
            });
        }

        [Test]
        public void OpenXmlTemplateExceptionIsATemplateException()
        {
            var settings = new ProcessSettings { UiCulture = German, ErrorMessages = CreateGermanMessages() };
            TemplateException e = OpenXmlTemplateException.Create(settings, TemplateErrorCode.ModelNotFound, "ds");

            Assert.Multiple(() =>
            {
                Assert.That(e, Is.InstanceOf<OpenXmlTemplateException>());
                Assert.That(e.ErrorCode, Is.EqualTo(TemplateErrorCode.ModelNotFound));
                Assert.That(e.Arguments, Is.EqualTo(new object[] { "ds" }));
                Assert.That(e.Message, Is.EqualTo("Modell ds nicht gefunden"));
                Assert.That(e.GetMessage(CultureInfo.InvariantCulture), Is.EqualTo("Model ds not found"));
            });
        }

        [Test]
        public void TemplateExceptionIsFormattedLikeTheDerivedOne()
        {
            var settings = new ProcessSettings { UiCulture = German, ErrorMessages = CreateGermanMessages() };
            var inner = new InvalidOperationException("boom");
            var e = TemplateException.Create(settings, inner, TemplateErrorCode.ModelNotFound, "ds");

            Assert.Multiple(() =>
            {
                Assert.That(e.GetType(), Is.EqualTo(typeof(TemplateException)));
                Assert.That(e.Message, Is.EqualTo("Modell ds nicht gefunden"));
                Assert.That(e.GetMessage(CultureInfo.InvariantCulture), Is.EqualTo("Model ds not found"));
                Assert.That(e.InnerException, Is.SameAs(inner));
            });
        }

        [Test]
        public void NestedTemplateExceptionIsTranslatedAsArgument()
        {
            var messages = CreateGermanMessages();
            var cause = TemplateException.Create(new ProcessSettings { ErrorMessages = messages }, TemplateErrorCode.ModelNotFound, "ds");
            var message = messages.Format(TemplateErrorCode.ErrorInCondition, German, cause, "ds.X > 1");
            Assert.That(message, Is.EqualTo("Modell ds nicht gefunden in condition 'ds.X > 1'"));
        }

        [Test]
        public void FreeTextMessageKeepsCodeNone()
        {
            var e = new TemplateException("free text");
            Assert.Multiple(() =>
            {
                Assert.That(e.ErrorCode, Is.EqualTo(TemplateErrorCode.None));
                Assert.That(e.Arguments, Is.Empty);
                Assert.That(e.GetMessage(German), Is.EqualTo("free text"));
            });
        }
    }
}
