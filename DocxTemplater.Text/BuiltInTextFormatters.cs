using System;
using DocxTemplater.Formatter;

namespace DocxTemplater.Text
{
    /// <summary><c>{{value}:format(pattern)}</c> / <c>{{value}:f(pattern)}</c> for <see cref="IFormattable"/> values.</summary>
    internal sealed class FormatPatternTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return (prefix.Equals("format", StringComparison.OrdinalIgnoreCase) || prefix.Equals("f", StringComparison.OrdinalIgnoreCase))
                   && typeof(IFormattable).IsAssignableFrom(type);
        }

        public TextValue ApplyFormat(ITextTemplateContext context, FormatterContext formatterContext)
        {
            var settings = context.Settings.ProcessSettings;
            if (formatterContext.Args.Length != 1)
            {
                throw TemplateException.Create(settings, TemplateErrorCode.FormatPatternRequiresArgument);
            }
            if (formatterContext.Value is not IFormattable formattable)
            {
                throw TemplateException.Create(settings, TemplateErrorCode.FormatterRequiresFormattable, formatterContext.Formatter, formatterContext.Placeholder);
            }
            var pattern = formatterContext.Args[0];
            try
            {
                return new TextValue(formattable.ToString(pattern, formatterContext.Culture));
            }
            catch (FormatException e)
            {
                throw TemplateException.Create(settings, e, TemplateErrorCode.FormatNotApplicable, pattern, formatterContext.Placeholder, formatterContext.Value.GetType());
            }
        }
    }

    /// <summary><c>{{value}:toupper()}</c> / <c>{{value}:tolower()}</c> for strings.</summary>
    internal sealed class CaseTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string)
                   && (prefix.Equals("toupper", StringComparison.OrdinalIgnoreCase) || prefix.Equals("tolower", StringComparison.OrdinalIgnoreCase));
        }

        public TextValue ApplyFormat(ITextTemplateContext context, FormatterContext formatterContext)
        {
            if (formatterContext.Value is not string text)
            {
                throw TemplateException.Create(context.Settings.ProcessSettings, TemplateErrorCode.FormatterRequiresString, formatterContext.Formatter, formatterContext.Placeholder);
            }
            return new TextValue(formatterContext.Formatter.Equals("toupper", StringComparison.OrdinalIgnoreCase)
                ? text.ToUpper(formatterContext.Culture)
                : text.ToLower(formatterContext.Culture));
        }
    }

    /// <summary><c>{{value}:html()}</c>: the string is trusted HTML and is inserted without encoding.</summary>
    internal sealed class HtmlTextFormatter : ITextFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string) && prefix.Equals("html", StringComparison.OrdinalIgnoreCase);
        }

        public TextValue ApplyFormat(ITextTemplateContext context, FormatterContext formatterContext)
        {
            return TextValue.Raw((string)formatterContext.Value);
        }
    }
}
