using System;
using System.Collections.Generic;
using System.Linq;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text
{
    /// <summary>
    /// Turns a resolved value into its text: the formatter named in the tag, or the default formatter of the value
    /// (<see cref="Model.ModelPropertyAttribute"/>), or <see cref="IFormattable"/> in the output culture.
    /// Mirrors the formatter selection of the Word renderer.
    /// </summary>
    internal sealed class TextValueFormatter
    {
        private readonly List<ITextFormatter> m_formatters = [new FormatPatternTextFormatter(), new HtmlTextFormatter(), new CaseTextFormatter()];

        public void Register(ITextFormatter formatter)
        {
            ArgumentNullException.ThrowIfNull(formatter);
            m_formatters.Add(formatter);
        }

        public TextValue Format(ITextTemplateContext context, PatternMatch match, ValueWithMetadata valueWithMetadata)
        {
            var value = valueWithMetadata.Value;
            if (value == null)
            {
                return new TextValue(string.Empty);
            }

            var settings = context.Settings;
            var formatterName = GetFormatter(match, valueWithMetadata, settings.ProcessSettings, out var arguments);
            if (!string.IsNullOrWhiteSpace(formatterName))
            {
                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), formatterName))
                    {
                        return formatter.ApplyFormat(context, new FormatterContext(match.Variable, formatterName, arguments, value, settings.Culture));
                    }
                }
            }

            return new TextValue(value is IFormattable formattable
                ? formattable.ToString(null, settings.Culture)
                : value.ToString() ?? string.Empty);
        }

        /// <summary>
        /// The formatter of the tag, or - if the tag has none - the default formatter of the value's metadata.
        /// </summary>
        private static string GetFormatter(PatternMatch match, ValueWithMetadata valueWithMetadata, ProcessSettings settings, out string[] arguments)
        {
            arguments = match.Arguments;
            var formatter = match.Formatter;
            if (string.IsNullOrWhiteSpace(formatter) && !string.IsNullOrWhiteSpace(valueWithMetadata.Metadata.DefaultFormatter))
            {
                var found = PatternMatcher.FindSyntaxPatterns("{{x}:" + valueWithMetadata.Metadata.DefaultFormatter + "}", settings).FirstOrDefault();
                if (found != null && !string.IsNullOrWhiteSpace(found.Formatter))
                {
                    formatter = found.Formatter;
                    arguments = found.Arguments;
                }
            }
            return formatter;
        }
    }
}
