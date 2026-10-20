namespace DocxTemplater.Text
{
    /// <summary>
    /// How the values of a <see cref="TextTemplate"/> are encoded in the output. The literal text of the template
    /// is never encoded - it is the author's markup.
    /// </summary>
    public enum TextEscaping
    {
        /// <summary>
        /// Values are inserted as they are, for plain text such as an e-mail subject. <c>{{:break}}</c> is a line feed.
        /// </summary>
        None,

        /// <summary>
        /// Values are HTML-encoded and their line breaks become <c>&lt;br/&gt;</c>, for HTML such as an e-mail body.
        /// A value that is trusted HTML is inserted unchanged with <c>{{value}:html()}</c>. <c>{{:break}}</c> is a <c>&lt;br/&gt;</c>.
        /// </summary>
        Html
    }
}
