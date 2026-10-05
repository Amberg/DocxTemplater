using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {

        /// <summary>
        /// Output culture of the document
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentUICulture;

        /// <summary>
        /// Culture of the user who generates the document. Determines the language of the error messages
        /// (exceptions, <see cref="TemplateSyntaxError.Message"/> and the errors written into the document with
        /// <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>), independent of the <see cref="Culture"/>
        /// of the document. A language is only used if it is available in <see cref="ErrorMessages"/>;
        /// otherwise the messages are English.
        /// default: <see cref="CultureInfo.CurrentUICulture"/>
        /// </summary>
        public CultureInfo UiCulture { get; set; } = CultureInfo.CurrentUICulture;

        /// <summary>
        /// The texts of the error messages. English is built in; other languages are added with
        /// <see cref="TemplateErrorMessages.AddLanguage"/>, e.g. from a language package.
        /// default: <see cref="TemplateErrorMessages.Default"/>, which is shared by all documents.
        /// </summary>
        public TemplateErrorMessages ErrorMessages { get; set; } = TemplateErrorMessages.Default;

        public BindingErrorHandling BindingErrorHandling { get; set; } = BindingErrorHandling.ThrowException;

        /// <summary>
        /// When enabled, this option removes leading or trailing newlines around template directives (e.g., {{#...}}, {{/}})
        /// from the final output. This allows templates to be more readable without affecting rendered formatting.
        /// default: false
        /// </summary>
        public bool IgnoreLineBreaksAroundTags { get; set; }

        /// <summary>
        /// When enabled, content controls whose tag is a placeholder (e.g. {{ds.Name}})
        /// are filled from the model. Default: false.
        /// </summary>
        public bool EnableContentControlTagBinding { get; set; }

        /// <summary>
        /// When enabled, the content of a sub-template made of a single top-level paragraph will be added to the
        /// destination paragraph instead of being inserted as a whole new paragraph. This allows to control the format
        /// of the sub document fragment via the target paragraph format (text alignment, etc.) and avoid to have an
        /// additional line return in the rendered document.
        /// This is especially useful for inline, short, templates
        /// default: false
        /// </summary>
        public bool InlineSubTemplates { get; set; }

        public static ProcessSettings Default => new();
    }
}
