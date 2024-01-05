using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Templates.Formatter
{
    internal class VariableReplacer
    {
        private readonly ModelDictionary m_models;
        private static readonly Regex FormatterRegex = new (@"(.+)(?:\((.*)?\))", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex m_variableRegex = new(@"\{\{([a-zA-Z0-9\.]+)\}(?::(\w+\(*\w*\)*))*\}", RegexOptions.Compiled);


        private readonly List<IFormatter> m_formatters;

        public VariableReplacer(ModelDictionary models)
        {
            m_models = models;
            m_formatters = new List<IFormatter>();
            m_formatters.Add(new FormatPatternFormatter());
            m_formatters.Add(new HtmlFormatter());
        }


        public void RegisterFormatter(IFormatter formatter)
        {
            m_formatters.Add(formatter);
        }

        /// <summary>
        /// the formatter string is the leading formatter prefix, e.g. "FORMAT" followed by the formatter arguments ae image(100,200)
        /// </summary>
        public void ApplyFormatter(string modelPath, object value, string formatterAsString, Text target)
        {
            if (value == null)
            {
                target.Text = string.Empty;
            }
            if (!string.IsNullOrWhiteSpace(formatterAsString))
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
                        formatter.ApplyFormat(modelPath, value, prefix, args, target);
                        return;
                    }
                }
            }
            target.Text = value.ToString();

        }

        public void ReplaceVariables(OpenXmlCompositeElement cloned)
        {
            foreach (var text in cloned.GetElementsWithMarker(ElementMarkers.Variable).OfType<Text>())
            {
               var variableMatch = m_variableRegex.Match(text.Text);
                if (!variableMatch.Success)
                {
                    throw new OpenXmlTemplateException($"Invalid variable syntax '{text.Text}'");
                }

                var variableName = variableMatch.Groups[1].Value;
                var formatterAsString = variableMatch.Groups[2].Value;

                var value = m_models.GetValue(variableName);
                ApplyFormatter(variableName, value, formatterAsString, text);
            }
        }
    }
}
