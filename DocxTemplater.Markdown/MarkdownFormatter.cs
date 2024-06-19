using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using Markdig;
using Markdig.Parsers;
using System;
using System.Linq;

namespace DocxTemplater.Markdown
{
    public class MarkdownFormatter : IFormatter, IFormatterInitialization
    {
        private readonly MarkDownFormatterConfiguration m_configuration;
        private IModelLookup m_modelLookup;
        private ProcessSettings m_processSettings;
        private IVariableReplacer m_variableReplacer;
        private IScriptCompiler m_scriptCompiler;
        private int m_nestingDepth;
        private MainDocumentPart m_mainDocumentPart;

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

            if (m_nestingDepth > 3)
            {
                throw new OpenXmlTemplateException("Markdown nesting depth exceeded");
            }

            m_nestingDepth++;
            var root = target.GetRoot();
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {
                var pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();
                var markdownDocument = MarkdownParser.Parse(mdText, pipeline);

                var parentParagraph = target.GetFirstAncestor<Paragraph>();
                // split the paragraph at the target
                var paragraph = (Paragraph)parentParagraph.SplitAfterElement(target).First();

                var renderer = new MarkdownToOpenXmlRenderer(paragraph, target, m_mainDocumentPart, m_configuration);
                var firstParagraph = renderer.CurrentParagraph;
                renderer.Render(markdownDocument);
                var lastParagraph = renderer.CurrentParagraph;
                try
                {
                    target.RemoveWithEmptyParent();
                    DoVariableReplacementInParagraphs(firstParagraph, lastParagraph);
                }
                catch (Exception e)
                {
                    throw new OpenXmlTemplateException("Variable Replacement in markdown failed", e);
                }
            }
            m_nestingDepth--;
            target.RemoveWithEmptyParent();
        }

        private void DoVariableReplacementInParagraphs(Paragraph firstParagraph, Paragraph lastParagraph)
        {
            var currentParagraph = firstParagraph;
            do
            {
                if (currentParagraph.InnerText.Contains('{'))
                {
                    var processor = new XmlNodeTemplate(currentParagraph, m_processSettings, m_modelLookup, m_variableReplacer, m_scriptCompiler, m_mainDocumentPart);
                    processor.Process();
                }
                currentParagraph = currentParagraph.NextSibling<Paragraph>();
                if (currentParagraph == null)
                {
                    break;
                }
            }
            while (currentParagraph != lastParagraph);
        }

        public void Initialize(IModelLookup modelLookup, IScriptCompiler scriptCompiler, IVariableReplacer variableReplacer,
            ProcessSettings processSettings, MainDocumentPart mainDocumentPart)
        {
            m_modelLookup = modelLookup;
            m_processSettings = processSettings;
            m_variableReplacer = variableReplacer;
            m_scriptCompiler = scriptCompiler;
            m_mainDocumentPart = mainDocumentPart;
        }
    }
}
