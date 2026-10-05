using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// Swiss German (<c>de-CH</c>) texts for the error messages of DocxTemplater.
    /// Register them once at startup: <c>TemplateErrorMessages.Default.AddSwissGerman();</c>
    /// </summary>
    /// <remarks>
    /// Swiss German spelling differs from standard German only in that it does not use the ß, so the texts are
    /// derived from <see cref="GermanErrorMessages"/> with ß replaced by ss. They are registered for <c>de-CH</c>
    /// only; register <see cref="GermanErrorMessages"/> as well if other German cultures (<c>de-DE</c>, <c>de-AT</c>)
    /// should get German messages.
    /// </remarks>
    public static class SwissGermanErrorMessages
    {
        /// <summary>The culture the texts are registered for.</summary>
        public static CultureInfo Culture { get; } = new("de-CH");

        /// <summary>
        /// Adds the Swiss German texts to <paramref name="messages"/>.
        /// </summary>
        /// <returns><paramref name="messages"/>, to chain calls.</returns>
        public static TemplateErrorMessages AddSwissGerman(this TemplateErrorMessages messages)
        {
            return messages.AddLanguage(Culture, Formats);
        }

        public static IReadOnlyDictionary<TemplateErrorCode, string> Formats { get; } = new ReadOnlyDictionary<TemplateErrorCode, string>(
            GermanErrorMessages.Formats.ToDictionary(x => x.Key, x => x.Value.Replace("ß", "ss")));
    }
}
