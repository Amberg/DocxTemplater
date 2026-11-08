using System;
using System.Collections.Generic;
using System.Globalization;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    /// <summary>
    /// The settings of a Word template. A <see cref="DocxTemplate"/> copies them when it is created, so changing
    /// them afterwards only affects templates created later; <see cref="DocxTemplate.Settings"/> is the copy of the
    /// template and can still be changed until it is processed.
    /// </summary>
    /// <remarks>
    /// All templates created with the same instance share a cache of the parsed conditions and expressions: create
    /// one instance per batch (e.g. a mailing or a print run). The cache has no size limit - it grows with the number
    /// of distinct expressions times the types of their values. Do not keep one instance for the lifetime of an
    /// application that renders user-edited templates.
    /// </remarks>
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
        /// The texts of the error messages. English is built in; installed language packages
        /// (<c>DocxTemplater.Localization.&lt;culture&gt;</c>) are loaded automatically, own texts are added with
        /// <see cref="TemplateErrorMessages.AddLanguage(CultureInfo, System.Collections.Generic.IReadOnlyDictionary{TemplateErrorCode, string})"/>.
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

        /// <summary>The formatters registered with <see cref="RegisterFormatter"/>, in registration order.</summary>
        internal IReadOnlyList<IFormatter> Formatters { get; private set; } = [];

        internal TagContentPreprocessors TagContentPreprocessors { get; private set; } = TagContentPreprocessors.None;

        /// <summary>The parsed conditions and expressions of all templates created with these settings.</summary>
        internal ScriptCache ScriptCache { get; } = new();

        public static ProcessSettings Default => new();

        /// <summary>
        /// Registers a formatter for all templates created with these settings. It is shared by the templates and must
        /// be thread-safe if they are processed in parallel.
        /// </summary>
        public void RegisterFormatter(IFormatter formatter)
        {
            ArgumentNullException.ThrowIfNull(formatter);
            Formatters = [.. Formatters, formatter];
        }

        /// <summary>
        /// Registers a preprocessor for the content of the tags (expressions, formatter arguments), applied before they
        /// are parsed, after the preprocessors registered before it. See <see cref="ITagContentPreprocessor"/>.
        /// </summary>
        public void RegisterTagContentPreprocessor(ITagContentPreprocessor preprocessor)
        {
            ArgumentNullException.ThrowIfNull(preprocessor);
            TagContentPreprocessors = TagContentPreprocessors.With(preprocessor);
        }

        /// <summary>
        /// The copy a template works with. The formatter and preprocessor lists are immutable and the cache is shared on
        /// purpose, so a shallow copy is enough.
        /// </summary>
        internal ProcessSettings Copy()
        {
            return (ProcessSettings)MemberwiseClone();
        }
    }
}
