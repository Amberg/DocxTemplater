using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    internal class FormatPatternFormatter : IFormatter
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
            if (prefix.Equals("FORMAT", StringComparison.CurrentCultureIgnoreCase) ||
                prefix.Equals("F", StringComparison.CurrentCultureIgnoreCase))
            {
                return type.IsAssignableTo(typeof(IFormattable));
            }

            return false;
        }

        internal static string Format(ProcessSettings settings, FormatterContext formatterContext)
        {
            if (formatterContext.Args.Length != 1)
            {
                throw OpenXmlTemplateException.Create(settings, TemplateErrorCode.FormatPatternRequiresArgument);
            }

            if (formatterContext.Value is IFormattable formattable)
            {
                var formatString = formatterContext.Args[0];
                try
                {
                    return formattable.ToString(formatString, formatterContext.Culture);
                }
                catch (FormatException e)
                {
                    throw OpenXmlTemplateException.Create(settings, e, TemplateErrorCode.FormatNotApplicable,
                        formatString, formatterContext.Placeholder, formatterContext.Value.GetType());
                }
            }

            throw OpenXmlTemplateException.Create(settings, TemplateErrorCode.FormatterRequiresFormattable,
                formatterContext.Formatter, formatterContext.Placeholder);
        }
    }
}
