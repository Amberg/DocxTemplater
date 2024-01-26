using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Formatter
{
    internal class VariableReplacer
    {
        private readonly ModelLookup m_models;
        private readonly ProcessSettings m_processSettings;
        private readonly List<IFormatter> m_formatters;

        public VariableReplacer(ModelLookup models, ProcessSettings processSettings)
        {
            m_models = models;
            m_processSettings = processSettings;
            m_formatters = new List<IFormatter>();
            m_formatters.Add(new FormatPatternFormatter());
            m_formatters.Add(new HtmlFormatter());
            m_formatters.Add(new CaseFormatter());
        }


        public void RegisterFormatter(IFormatter formatter)
        {
            m_formatters.Add(formatter);
        }

        /// <summary>
        /// the formatter string is the leading formatter prefix, e.g. "FORMAT" followed by the formatter arguments ae image(100,200)
        /// </summary>
        public void ApplyFormatter(PatternMatch patternMatch, object value, Text target)
        {
            if (value == null)
            {
                target.Text = string.Empty;
                return;
            }
            if (!string.IsNullOrWhiteSpace(patternMatch.Formatter))
            {
                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), patternMatch.Formatter))
                    {
                        var context = new FormatterContext(patternMatch.Variable, patternMatch.Formatter, patternMatch.Arguments, value, m_processSettings.Culture);
                        formatter.ApplyFormat(context, target);
                        return;
                    }
                }
            }
            if (value is IFormattable formattable)
            {
                target.Text = formattable.ToString(null, m_processSettings.Culture);
                return;
            }
            target.Text = value.ToString() ?? string.Empty;

        }

        public void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content)
        {
            foreach (var element in content)
            {
                ReplaceVariables(element);
            }
        }

        public void ReplaceVariables(OpenXmlElement cloned)
        {
            var variables = cloned.GetElementsWithMarker(PatternType.Variable).OfType<Text>().ToList();
            foreach (var text in variables)
            {
                var variableMatch = PatternMatcher.FindSyntaxPatterns(text.Text).FirstOrDefault() ?? throw new OpenXmlTemplateException($"Invalid variable syntax '{text.Text}'");
                try
                {
                    var value = m_models.GetValue(variableMatch.Variable);
                    ApplyFormatter(variableMatch, value, text);
                }
                catch (OpenXmlTemplateException) when (m_processSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
                {
                    if (m_processSettings.BindingErrorHandling == BindingErrorHandling.SkipBindingAndRemoveContent)
                    {
                        text.RemoveWithEmptyParent();
                    }
                }
            }
        }
    }
}
