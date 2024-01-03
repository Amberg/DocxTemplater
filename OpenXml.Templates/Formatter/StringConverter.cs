using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OpenXml.Templates.Formatter
{
    internal class StringConverter
    {
        private static readonly Regex FormatterRegex = new (@"(.+)(?:\((.*)?\))", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly List<IFormatter> m_formatters;

        public StringConverter()
        {
            m_formatters = new List<IFormatter>();
            m_formatters.Add(new DateTimeFormatter());
        }

        /// <summary>
        /// the formatter string is the leading formatter prefix, e.g. "FORMAT" followed by the formatter arguments ae image(100,200)
        /// </summary>
        public void ApplyFormatter(object value, string formatterAsString, Text target)
        {
            if (value == null)
            {
                target.Text = string.Empty;
            }
            if (formatterAsString != null)
            {
                var matchResult = FormatterRegex.Match(formatterAsString);
                if (!matchResult.Success)
                {
                    throw new OpenXmlTemplateException($"Unknown formatter '{formatterAsString}'");
                }

                var prefix = matchResult.Groups[1].Value;
                var args = matchResult.Groups[2].Value.Split(',');

                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), prefix))
                    {
                        target.Text = formatter.Format(value, prefix, args);
                        return;
                    }
                }
            }
            target.Text = value.ToString();

        }
    }
}
