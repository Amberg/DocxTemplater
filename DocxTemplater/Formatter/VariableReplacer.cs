﻿using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Formatter
{
    internal class VariableReplacer : IVariableReplacer
    {
        private readonly IModelLookup m_models;
        private readonly ProcessSettings m_processSettings;
        private readonly List<IFormatter> m_formatters;

        public VariableReplacer(IModelLookup models, ProcessSettings processSettings)
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
        private void ApplyFormatter(PatternMatch patternMatch, ValueWithMetadata valueWithMetadata, Text target)
        {
            var value = valueWithMetadata.Value;
            if (value == null)
            {
                target.Text = string.Empty;
                return;
            }
            var formatterText = GetFormatterText(patternMatch, valueWithMetadata, out string[] formaterArguments);
            if (!string.IsNullOrWhiteSpace(formatterText))
            {
                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), formatterText))
                    {
                        var context = new FormatterContext(patternMatch.Variable, formatterText, formaterArguments, value, m_processSettings.Culture);
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
                    var valueWithMetadata = m_models.GetValueWithMetadata(variableMatch.Variable);
                    ApplyFormatter(variableMatch, valueWithMetadata, text);
                    VariableReplacer.SplitNewLinesInText(text);
                }
                catch (Exception e) when (e is OpenXmlTemplateException or FormatException)
                {
                    if (m_processSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
                    {
                        text.RemoveWithEmptyParent();
                    }
                    else
                    {
                        throw new OpenXmlTemplateException($"'{text.InnerText}' could not be replaced: {text.ElementBeforeInDocument<Text>()?.InnerText} >> {text.InnerText} << {text.ElementAfterInDocument<Text>()?.InnerText}", e);
                    }
                }
            }
        }


        /// <summary>
        /// Use the formatter from the template, if not available use the default formatter from the metadata
        /// set through <see cref="ModelPropertyAttribute"/>
        /// </summary>
        private static string GetFormatterText(PatternMatch patternMatch, ValueWithMetadata valueWithMetadata, out string[] formatterArguments)
        {
            formatterArguments = patternMatch.Arguments;
            var formatterText = patternMatch.Formatter;
            if (string.IsNullOrWhiteSpace(formatterText))
            {
                if (!string.IsNullOrWhiteSpace(valueWithMetadata.Metadata.DefaultFormatter))
                {
                    // try to parse default formatter from metadata
                    var found = PatternMatcher.FindSyntaxPatterns("{{x}:" + valueWithMetadata.Metadata.DefaultFormatter + "}").FirstOrDefault();
                    if (found != null && !string.IsNullOrWhiteSpace(found.Formatter))
                    {
                        formatterText = found.Formatter;
                        formatterArguments = found.Arguments;
                    }
                }
            }
            return formatterText;
        }

        /// <summary>
        /// Insert Breaks for line breaks in the text
        /// </summary>
        private static void SplitNewLinesInText(Text text)
        {
            if (text.Parent == null)
            {
                return;
            }
            text.Text = text.Text.Replace("\r\n", "\n").Replace("\r", "\n");
            if (text.Text.Contains('\n'))
            {
                var parts = text.Text.Split('\n');
                OpenXmlElement lastElement = text;
                foreach (var part in parts)
                {
                    lastElement = lastElement.InsertAfterSelf(new Text(part));
                    lastElement = lastElement.InsertAfterSelf(new Break());
                }

                text.Remove();
            }
        }
    }
}
