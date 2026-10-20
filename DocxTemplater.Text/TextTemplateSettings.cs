using System.Globalization;

namespace DocxTemplater.Text
{
    /// <summary>
    /// The settings of a <see cref="TextTemplate"/>: the part shared with Word templates (culture, error handling,
    /// error messages) plus the <see cref="Escaping"/> of the values.
    /// </summary>
    public sealed class TextTemplateSettings
    {
        /// <summary>
        /// Culture the values are formatted with (dates, numbers, <c>ToUpper</c>).
        /// default: <see cref="CultureInfo.CurrentUICulture"/>
        /// </summary>
        public CultureInfo Culture
        {
            get => ProcessSettings.Culture;
            set => ProcessSettings.Culture = value;
        }

        /// <summary>
        /// Culture of the user who renders the template. Determines the language of the error messages (exceptions,
        /// <see cref="TemplateSyntaxError.Message"/> and the errors written into the output with
        /// <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>), independent of the <see cref="Culture"/>
        /// of the output. A language is only used if it is available in <see cref="ErrorMessages"/>; otherwise the
        /// messages are English.
        /// default: <see cref="CultureInfo.CurrentUICulture"/>
        /// </summary>
        public CultureInfo UiCulture
        {
            get => ProcessSettings.UiCulture;
            set => ProcessSettings.UiCulture = value;
        }

        /// <summary>
        /// The texts of the error messages, see <see cref="ProcessSettings.ErrorMessages"/>.
        /// default: <see cref="TemplateErrorMessages.Default"/>
        /// </summary>
        public TemplateErrorMessages ErrorMessages
        {
            get => ProcessSettings.ErrorMessages;
            set => ProcessSettings.ErrorMessages = value;
        }

        /// <summary>
        /// What happens when a placeholder cannot be bound to the model. With
        /// <see cref="BindingErrorHandling.HighlightErrorsInDocument"/> the failed tags stay in the output
        /// (marked red in <see cref="TextEscaping.Html"/>) and all errors are listed at the top of the output
        /// and in <see cref="TextTemplate.Errors"/>.
        /// default: <see cref="BindingErrorHandling.ThrowException"/>
        /// </summary>
        public BindingErrorHandling BindingErrorHandling
        {
            get => ProcessSettings.BindingErrorHandling;
            set => ProcessSettings.BindingErrorHandling = value;
        }

        /// <summary>
        /// When enabled, a line that holds nothing but block tags (<c>{{#Items}}</c>, <c>{{/}}</c>, <c>{{:}}</c>,
        /// <c>{{:s:}}</c>, <c>{{#case: ...}}</c>, ...) and whitespace is removed from the output together with its
        /// line break, so a template can be indented and spread over lines without affecting the output.
        /// Lines with values or <c>{{:break}}</c> are not affected.
        /// default: false
        /// </summary>
        public bool IgnoreLineBreaksAroundTags
        {
            get => ProcessSettings.IgnoreLineBreaksAroundTags;
            set => ProcessSettings.IgnoreLineBreaksAroundTags = value;
        }

        /// <summary>
        /// How values are encoded in the output. The default is <see cref="TextEscaping.Html"/>, so a value cannot
        /// break the markup of an HTML template unless it is inserted raw on purpose; use <see cref="TextEscaping.None"/>
        /// for plain text.
        /// </summary>
        public TextEscaping Escaping { get; set; } = TextEscaping.Html;

        /// <summary>
        /// The settings as seen by the core (model lookup, expressions, error messages). The properties above write
        /// through to this instance.
        /// </summary>
        internal ProcessSettings ProcessSettings { get; } = new();

        public static TextTemplateSettings Default => new();
    }
}
