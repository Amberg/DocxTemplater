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
                return type.IsAssignableTo(typeof(IFormattable));
            }

            return false;
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {

            if (formatterContext.Args.Length != 1)
            {
                throw OpenXmlTemplateException.Create(templateContext.ProcessSettings, TemplateErrorCode.FormatPatternRequiresArgument);
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
                    throw OpenXmlTemplateException.Create(templateContext.ProcessSettings, e, TemplateErrorCode.FormatNotApplicable,
                        formatString, formatterContext.Placeholder, formatterContext.Value.GetType());
                }
            }
            else
            {
                throw OpenXmlTemplateException.Create(templateContext.ProcessSettings, TemplateErrorCode.FormatterRequiresFormattable,
                    formatterContext.Formatter, formatterContext.Placeholder);
            }
        }
    }
}
