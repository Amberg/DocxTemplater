using System.Globalization;

namespace DocxTemplater.Formatter
{
    public class FormatterContext
    {
        public FormatterContext(string placeholder, string formatter, string[] args, object value, CultureInfo culture)
        {
            Placeholder = placeholder;
            Formatter = formatter;
            Args = args;
            Culture = culture;
            Value = value;
        }

        public object Value { get; }

        public string[] Args { get; }

        public CultureInfo Culture { get; }

        /// <summary>
        ///  Formatter without parentheses
        /// </summary>
        public string Formatter { get; }

        public string Placeholder { get; }
    }
}
