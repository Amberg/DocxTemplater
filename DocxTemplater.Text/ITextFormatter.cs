using System;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text
{
    /// <summary>
    /// The output of an <see cref="ITextFormatter"/>.
    /// </summary>
    /// <param name="Text">The formatted value.</param>
    /// <param name="IsRaw">
    /// <c>true</c> if the text is already in the output format (e.g. trusted HTML) and must be inserted as it is,
    /// <c>false</c> if it is a plain value that is encoded according to <see cref="TextTemplateSettings.Escaping"/>.
    /// </param>
    public readonly record struct TextValue(string Text, bool IsRaw)
    {
        /// <summary>A plain value that is encoded according to <see cref="TextTemplateSettings.Escaping"/>.</summary>
        public TextValue(string text) : this(text, false)
        {
        }

        /// <summary>A value that is already in the output format and is inserted unchanged.</summary>
        public static TextValue Raw(string text)
        {
            return new TextValue(text, true);
        }
    }

    /// <summary>
    /// Formats the value of a placeholder of a <see cref="TextTemplate"/>, selected by the formatter name in the
    /// template (<c>{{value}:name(args)}</c>). Register own formatters with
    /// <see cref="TextTemplateSettings.RegisterFormatter"/>; the built-in ones are <c>format</c> / <c>f</c>,
    /// <c>toupper</c> / <c>tolower</c> and <c>html</c>. <c>default('text')</c> / <c>ifempty('text')</c> is handled before
    /// the formatters, since they never see a <c>null</c> value; a formatter of this name is never called.
    /// </summary>
    public interface ITextFormatter
    {
        /// <param name="type">The type of the value.</param>
        /// <param name="prefix">The formatter name without arguments, as written in the template.</param>
        bool CanHandle(Type type, string prefix);

        /// <summary>
        /// Formats <see cref="FormatterContext.Value"/>. An exception - preferably a <see cref="TemplateException"/> for a
        /// value or argument the formatter cannot handle - is reported according to <see cref="TextTemplateSettings.BindingErrorHandling"/>.
        /// </summary>
        /// <param name="settings">The settings of the template, e.g. for the language of an error message.</param>
        TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext);
    }
}
