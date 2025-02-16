using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    public interface IFormatter
    {
        bool CanHandle(Type type, string prefix);

        void ApplyFormat(TemplateProcessingContext templateContext, FormatterContext formatterContext, Text target);
    }
}
