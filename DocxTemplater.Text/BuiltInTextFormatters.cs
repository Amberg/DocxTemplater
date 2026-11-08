using System;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text
{
    /// <summary><c>{{value}:format(pattern)}</c> / <c>{{value}:f(pattern)}</c> for <see cref="IFormattable"/> values.</summary>
    internal sealed class FormatPatternTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return FormatPatternFormatter.Handles(type, prefix);
        }

        public TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext)
        {
            return new TextValue(FormatPatternFormatter.Format(settings.ProcessSettings, formatterContext));
        }
    }

    /// <summary><c>{{value}:toupper()}</c> / <c>{{value}:tolower()}</c> for strings.</summary>
    internal sealed class CaseTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return CaseFormatter.Handles(type, prefix);
        }

        public TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext)
        {
            return new TextValue(CaseFormatter.Format(settings.ProcessSettings, formatterContext));
        }
    }

    /// <summary><c>{{value}:html()}</c>: the string is trusted HTML and is inserted without encoding.</summary>
    internal sealed class HtmlTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string) && prefix.Equals("html", StringComparison.OrdinalIgnoreCase);
        }

        public TextValue ApplyFormat(TextTemplateSettings settings, FormatterContext formatterContext)
        {
            return TextValue.Raw((string)formatterContext.Value);
        }
    }
}
