using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    internal class FormatPatternFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            if (prefix.Equals("FORMAT", StringComparison.CurrentCultureIgnoreCase) || prefix.Equals("F", StringComparison.CurrentCultureIgnoreCase))
            {
                return type.IsAssignableTo(typeof(IFormattable));
            }
            return false;
        }

        public void ApplyFormat(FormatterContext context, Text target)
        {
            if (context.Args.Length != 1)
            {
                throw new OpenXmlTemplateException($"DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)");
            }
            if (context.Value is IFormattable formattable)
            {
                target.Text = formattable.ToString(context.Args[0], null);
            }
            else
            {
                throw new OpenXmlTemplateException($"Formatter {context.Formatter} can only be applied to IFormattable objects - property {context.Placeholder}");
            }
        }
    }

}
