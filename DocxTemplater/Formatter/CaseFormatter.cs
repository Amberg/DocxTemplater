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

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Value is string str)
            {
                if (formatterContext.Formatter.Equals("toupper", StringComparison.InvariantCultureIgnoreCase))
                {
                    target.Text = str.ToUpper(formatterContext.Culture);
                }
                else if (formatterContext.Formatter.Equals("tolower", StringComparison.InvariantCultureIgnoreCase))
                {
                    target.Text = str.ToLower(formatterContext.Culture);
                }
            }
            else
            {
                throw new OpenXmlTemplateException($"Formatter {formatterContext.Formatter} can only be applied to string objects - property {formatterContext.Placeholder}");
            }
        }
    }
}
