using System.Collections.Generic;
using System.Globalization;

namespace DocxTemplater.Localization
{
    /// <summary>
    /// The texts of the error messages in one language, see <see cref="TemplateErrorMessages"/>.
    /// </summary>
    /// <remarks>
    /// Language packages are NuGet packages named <c>DocxTemplater.Localization.&lt;language tag&gt;</c>
    /// (e.g. <c>DocxTemplater.Localization.de-CH</c>, the tag being <see cref="CultureInfo.Name"/>) whose assembly
    /// contains a public implementation of this interface with a parameterless constructor.
    /// <see cref="TemplateErrorMessages"/> loads such a package automatically the first time a message is formatted
    /// for its culture (<see cref="TemplateErrorMessages.AutoLoadLanguagePackages"/>). A pack can also be registered
    /// explicitly with <see cref="TemplateErrorMessages.AddLanguage(ITemplateLanguagePack)"/>.
    /// </remarks>
    public interface ITemplateLanguagePack
    {
        /// <summary>
        /// The culture the texts are written for, neutral (<c>de</c>) or specific (<c>de-CH</c>).
        /// </summary>
        CultureInfo Culture { get; }

        /// <summary>
        /// The composite format string per code. The pack does not have to be complete; missing codes fall back to the
        /// parent culture and English.
        /// </summary>
        IReadOnlyDictionary<TemplateErrorCode, string> Formats { get; }
    }
}
