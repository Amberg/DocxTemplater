using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace DocxTemplater.Formatter
{
    internal class CaseFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            if (type == typeof(string))
            {
                return prefix.Equals("toupper", StringComparison.InvariantCultureIgnoreCase) || prefix.Equals("tolower", StringComparison.InvariantCultureIgnoreCase);
            }
            return false;
        }

        public void ApplyFormat(FormatterContext context, Text target)
        {
            if (context.Value is string str)
            {
                if (context.Formatter.Equals("toupper", StringComparison.InvariantCultureIgnoreCase))
                {
                    target.Text = str.ToUpper();
                }
                else if (context.Formatter.Equals("tolower", StringComparison.InvariantCultureIgnoreCase))
                {
                    target.Text = str.ToLower();
                }
            }
            else
            {
                throw new OpenXmlTemplateException($"Formatter {context.Formatter} can only be applied to string objects - property {context.Placeholder}");
            }
        }
    }
}
