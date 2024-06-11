using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using Markdig.Parsers;
using System;
using DocumentFormat.OpenXml.Packaging;
using Markdig;

namespace DocxTemplater.Markdown
{
    public class MarkdownFormatter : IFormatter
    {
        private readonly MarkDownFormatterConfiguration m_configuration;

        public MarkdownFormatter(MarkDownFormatterConfiguration configuration = null)
        {
            m_configuration = configuration ?? MarkDownFormatterConfiguration.Default;
        }

        public bool CanHandle(Type type, string prefix)
        {
            string prefixUpper = prefix.ToUpper();
            return prefixUpper is "MD" && type == typeof(string);
        }

        public void ApplyFormat(FormatterContext context, Text target)
        {
            if (context.Value is not string mdText)
            {
                return;
            }

            var root = target.GetRoot();
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {

                if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                {
                    var pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();
                    var markdownDocument = MarkdownParser.Parse(mdText, pipeline);
                    var renderer = new MarkdownToOpenXmlRenderer(target, mainDocumentPart, m_configuration);
                    renderer.Render(markdownDocument);
                }
                else
                {
                    throw new InvalidOperationException("Markdown currently only supported in MainDocument");
                }
            }
            target.RemoveWithEmptyParent();
        }
    }
}
