namespace DocxTemplater.Text
{
    /// <summary>
    /// A binding error of a <see cref="TextTemplate.Process"/>: a
    /// placeholder, block or expression that could not be rendered with the given models.
    /// </summary>
    public sealed class TextRenderError
    {
        internal TextRenderError(string tag, int index, int length, TemplateException exception)
        {
            Tag = tag;
            Index = index;
            Length = length;
            Exception = exception;
        }

        /// <summary>The tag that could not be rendered, e.g. <c>{{recipient.Name}}</c> or <c>{?{ds.Count &gt; 1}}</c>.</summary>
        public string Tag { get; }

        /// <summary>Position of the <see cref="Tag"/> in the template text.</summary>
        public int Index { get; }

        /// <summary>Length of the <see cref="Tag"/>; 0 if the error has no location.</summary>
        public int Length { get; }

        /// <summary>
        /// The error, with its message in the <see cref="TextTemplateSettings.UiCulture"/>, its
        /// <see cref="TemplateException.ErrorCode"/> and the cause as inner exception, if any.
        /// </summary>
        public TemplateException Exception { get; }

        public override string ToString()
        {
            return Exception.Message;
        }
    }
}
