using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{

    /// <summary>
    /// Interface for formatters
    /// Formatters are used to apply custom formatting to the output of a placeholder.
    /// </summary>
    public interface IFormatter
    {
        public bool CanHandle(Type type, string prefix);

        void ApplyFormat(FormatterContext context, Text target);
    }
}
