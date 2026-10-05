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
        internal TemplateSyntaxError(TemplateSyntaxErrorSeverity severity, string part, string tag, string message, string context)
        {
            Severity = severity;
            Part = part;
            Tag = tag;
            Message = message;
            Context = context;
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
        /// Description of the error.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Text surrounding the error, to help locating it in the document.
        /// </summary>
        public string Context { get; }

        public override string ToString()
        {
            return $"{Part}: {Message} (near '{Context}')";
        }
    }
}
