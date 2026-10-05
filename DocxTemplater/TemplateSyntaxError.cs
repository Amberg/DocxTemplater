using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocxTemplater
{
    public enum TemplateSyntaxErrorSeverity
    {
        /// <summary>
        /// The template can still be rendered, but the construct is most likely not what the author intended
        /// (e.g. a malformed tag that stays in the document as plain text).
        /// </summary>
        Warning,

        /// <summary>
        /// The template cannot be rendered; <see cref="DocxTemplate.Process"/> throws an <see cref="OpenXmlTemplateException"/>.
        /// </summary>
        Error
    }

    /// <summary>
    /// A syntax error found by <see cref="DocxTemplate.ValidateTemplateSyntax"/>.
    /// </summary>
    public sealed class TemplateSyntaxError
    {
        private readonly TemplateErrorMessages m_messages;
        private readonly CultureInfo m_culture;

        /// <param name="index">Position of the <paramref name="tag"/> in the text of the part.</param>
        /// <param name="length">Length of the <paramref name="tag"/> in the text of the part.</param>
        /// <param name="message">The free text for an error without code (<see cref="TemplateErrorCode.None"/>); otherwise <c>null</c>.</param>
        internal TemplateSyntaxError(TemplateSyntaxErrorSeverity severity, string part, string tag, string context,
            TemplateErrorCode errorCode, IReadOnlyList<object> arguments, TemplateErrorMessages messages, CultureInfo culture,
            int index = 0, int length = 0, string message = null)
        {
            Severity = severity;
            Part = part;
            Tag = tag;
            Context = context;
            Index = index;
            Length = length;
            ErrorCode = errorCode;
            Arguments = arguments ?? Array.Empty<object>();
            m_messages = messages ?? TemplateErrorMessages.Default;
            m_culture = culture ?? CultureInfo.InvariantCulture;
            Message = errorCode == TemplateErrorCode.None ? message ?? string.Empty : m_messages.Format(errorCode, m_culture, Arguments);
        }

        public TemplateSyntaxErrorSeverity Severity { get; }

        /// <summary>
        /// The document part containing the error: "Body", "Header" or "Footer".
        /// </summary>
        public string Part { get; }

        /// <summary>
        /// The offending template text, e.g. <c>{{#Items}}</c> or a malformed fragment like <c>{{Name}</c>.
        /// </summary>
        public string Tag { get; }

        /// <summary>
        /// Description of the error in the <see cref="ProcessSettings.UiCulture"/>.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Text surrounding the error, to help locating it in the document.
        /// </summary>
        public string Context { get; }

        /// <summary>
        /// Position of the <see cref="Tag"/> in the text of the part, as seen by the parser (all text runs concatenated).
        /// </summary>
        internal int Index { get; }

        /// <summary>
        /// Length of the <see cref="Tag"/> in the text of the part; 0 if the error has no location.
        /// </summary>
        internal int Length { get; }

        /// <summary>
        /// Identifies the error independent of the language of the <see cref="Message"/>.
        /// </summary>
        public TemplateErrorCode ErrorCode { get; }

        /// <summary>
        /// The arguments of the message, see the documentation of the <see cref="ErrorCode"/>'s value for their meaning.
        /// </summary>
        public IReadOnlyList<object> Arguments { get; }

        /// <summary>
        /// The <see cref="Message"/> in another language.
        /// </summary>
        public string GetMessage(CultureInfo culture)
        {
            return GetMessage(m_messages, culture);
        }

        /// <inheritdoc cref="GetMessage(CultureInfo)"/>
        public string GetMessage(TemplateErrorMessages messages, CultureInfo culture)
        {
            ArgumentNullException.ThrowIfNull(messages);
            return ErrorCode == TemplateErrorCode.None ? Message : messages.Format(ErrorCode, culture, Arguments);
        }

        /// <summary>
        /// The error with its location, e.g. <c>Body: '{{/Orders}}' does not match '{{#Items}}' (near '...')</c>.
        /// </summary>
        public override string ToString()
        {
            return ToString(m_messages, m_culture);
        }

        /// <inheritdoc cref="ToString()"/>
        public string ToString(CultureInfo culture)
        {
            return ToString(m_messages, culture);
        }

        /// <inheritdoc cref="ToString()"/>
        public string ToString(TemplateErrorMessages messages, CultureInfo culture)
        {
            ArgumentNullException.ThrowIfNull(messages);
            return messages.Format(TemplateErrorCode.SyntaxErrorLocation, culture, Part, GetMessage(messages, culture), Context);
        }
    }
}
