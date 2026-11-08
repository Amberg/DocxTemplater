using System.Collections.Generic;

namespace DocxTemplater.Text
{
    /// <summary>
    /// The output of one <see cref="TextTemplate.Process"/>: the text
    /// and the binding errors that occurred while rendering it.
    /// </summary>
    public sealed class TextRenderResult
    {
        internal TextRenderResult(string text, IReadOnlyList<TextRenderError> errors)
        {
            Text = text;
            Errors = errors;
        }

        /// <summary>The rendered text.</summary>
        public string Text { get; }

        /// <summary>
        /// The binding errors in order of occurrence; empty if the template rendered without errors. With
        /// <see cref="BindingErrorHandling.SkipBindingAndRemoveContent"/> the failed tags are left out of
        /// <see cref="Text"/> and are reported only here.
        /// </summary>
        public IReadOnlyList<TextRenderError> Errors { get; }

        public bool HasErrors => Errors.Count > 0;

        public override string ToString()
        {
            return Text;
        }
    }
}
