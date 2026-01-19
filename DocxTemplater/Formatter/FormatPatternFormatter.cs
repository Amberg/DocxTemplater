using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    internal class FormatPatternFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            if (prefix.Equals("FORMAT", StringComparison.CurrentCultureIgnoreCase) ||
                prefix.Equals("F", StringComparison.CurrentCultureIgnoreCase))
            {
                return typeof(IFormattable).IsAssignableFrom(type);
            }

            return false;
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {

            if (formatterContext.Args.Length != 1)
            {
                throw new OpenXmlTemplateException(
                    $"DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)");
            }

            if (formatterContext.Value is IFormattable formattable)
            {
                var formatString = formatterContext.Args[0];
                try
                {
                    target.Text = formattable.ToString(formatString, formatterContext.Culture);
                }
                catch (FormatException e)
                {
                    throw new OpenXmlTemplateException($"Format {formatString} cannot be applied to {formatterContext.Placeholder} of type {formatterContext.Value.GetType()}", e);
                }
            }
            else
            {
                throw new OpenXmlTemplateException(
                    $"Formatter {formatterContext.Formatter} can only be applied to IFormattable objects - property {formatterContext.Placeholder}");
            }
        }
    }
}
