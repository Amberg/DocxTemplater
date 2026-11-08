using System;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text
{
    /// <summary>
    /// Turns a resolved value into its text: the formatter named in the tag, or the default formatter of the value
    /// (<see cref="Model.ModelPropertyAttribute"/>), or <see cref="IFormattable"/> in the output culture.
    /// Mirrors the formatter selection of the Word renderer.
    /// </summary>
    /// <remarks>Immutable, so it can be shared by concurrent renders; <see cref="With"/> adds a formatter.</remarks>
    internal sealed class TextValueFormatter
    {
        public static readonly TextValueFormatter BuiltIn = new([new FormatPatternTextFormatter(), new HtmlTextFormatter(), new CaseTextFormatter()]);

        private readonly ITextFormatter[] m_formatters;

        private TextValueFormatter(ITextFormatter[] formatters)
        {
            m_formatters = formatters;
        }

        public TextValueFormatter With(ITextFormatter formatter)
        {
            ArgumentNullException.ThrowIfNull(formatter);
            return new TextValueFormatter([.. m_formatters, formatter]);
        }

        /// <exception cref="TemplateException">No formatter handles the formatter of the tag (<see cref="TemplateErrorCode.UnknownFormatter"/>).</exception>
        public TextValue Format(TextTemplateSettings settings, PatternMatch match, ValueWithMetadata valueWithMetadata)
        {
            var value = valueWithMetadata.Value;
            if (DefaultValue.TryGetReplacement(ref match, value, out var replacement))
            {
                return new TextValue(replacement);
            }
            if (value == null)
            {
                return new TextValue(string.Empty);
            }

            var formatterName = VariableReplacer.GetFormatterText(match, valueWithMetadata, settings.ProcessSettings, out var arguments);
            if (!string.IsNullOrWhiteSpace(formatterName))
            {
                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), formatterName))
                    {
                        return formatter.ApplyFormat(settings, new FormatterContext(match.Variable, formatterName, arguments, value, settings.Culture));
                    }
                }
                // a default formatter of the model may be meant for Word documents only - only the tag's own formatter is checked
                if (!string.IsNullOrWhiteSpace(match.Formatter))
                {
                    throw TemplateException.Create(settings.ProcessSettings, TemplateErrorCode.UnknownFormatter, match.Match.Value, formatterName, value.GetType());
                }
            }

            return new TextValue(value is IFormattable formattable
                ? formattable.ToString(null, settings.Culture)
                : value.ToString() ?? string.Empty);
        }
    }
}
