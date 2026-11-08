using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace DocxTemplater.Formatter
{
    internal class CaseFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return Handles(type, prefix);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            target.Text = Format(templateContext.ProcessSettings, formatterContext);
        }

        // shared with the text templates
        internal static bool Handles(Type type, string prefix)
        {
            if (type == typeof(string))
            {
                return prefix.Equals("toupper", StringComparison.InvariantCultureIgnoreCase) || prefix.Equals("tolower", StringComparison.InvariantCultureIgnoreCase);
            }
            return false;
        }

        internal static string Format(ProcessSettings settings, FormatterContext formatterContext)
        {
            if (formatterContext.Value is not string str)
            {
                throw OpenXmlTemplateException.Create(settings, TemplateErrorCode.FormatterRequiresString,
                    formatterContext.Formatter, formatterContext.Placeholder);
            }
            return formatterContext.Formatter.Equals("toupper", StringComparison.InvariantCultureIgnoreCase)
                ? str.ToUpper(formatterContext.Culture)
                : str.ToLower(formatterContext.Culture);
        }
    }
}
