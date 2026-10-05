using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// Swiss German texts for the error messages of DocxTemplater, culture <c>de-CH</c>.
    /// Loaded automatically by <see cref="TemplateErrorMessages"/> when this package is installed.
    /// </summary>
    /// <remarks>
    /// Swiss German spelling differs from standard German only in that it does not use the ß, so the texts are
    /// derived from <see cref="GermanErrorMessages"/> with ß replaced by ss. They apply to <c>de-CH</c> only; other
    /// German cultures (<c>de-DE</c>, <c>de-AT</c>) use the <c>DocxTemplater.Localization.de</c> package this one
    /// depends on.
    /// </remarks>
    public sealed class SwissGermanErrorMessages : ITemplateLanguagePack
    {
        /// <inheritdoc/>
        public CultureInfo Culture { get; } = new("de-CH");

        /// <inheritdoc/>
        public IReadOnlyDictionary<TemplateErrorCode, string> Formats => Texts;

        private static readonly ReadOnlyDictionary<TemplateErrorCode, string> Texts = new(
            new GermanErrorMessages().Formats.ToDictionary(x => x.Key, x => x.Value.Replace("ß", "ss")));
    }
}
