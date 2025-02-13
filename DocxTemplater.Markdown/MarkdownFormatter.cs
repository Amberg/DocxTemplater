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
            try
            {
                var root = target.GetRoot();
                if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
                {
                    var pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();
                    var markdownDocument = MarkdownParser.Parse(mdText, pipeline);


                    // split the paragraph at the target
                    var split = target.GetFirstAncestor<Paragraph>().SplitAfterElement(target);
                    Paragraph paragraphBeforeMd = split.OfType<Paragraph>().First();
                    Paragraph paragraphAfterMd = split.OfType<Paragraph>().Last();

                    // create a container for the markdown as render target
                    var renderedMarkdownContainer = new Body();
                    var containerParagraph = new Paragraph();
                    renderedMarkdownContainer.Append(containerParagraph);

                    var renderer = new MarkdownToOpenXmlRenderer(containerParagraph, target, m_mainDocumentPart, m_configuration);
                    renderer.Render(markdownDocument);
                    try
                    {
                        DoVariableReplacementInParagraphs(renderedMarkdownContainer);
                    }
                    catch (Exception e)
                    {
                        throw new OpenXmlTemplateException("Variable Replacement in markdown failed", e);
                    }
                    containerParagraph.RemoveWithEmptyParent();

#if DEBUG
                    Console.WriteLine("----------- Markdown Converted to OpenXML --------");
                    Console.WriteLine(renderedMarkdownContainer.ToPrettyPrintXml());
#endif

                    var convertedMdElements = renderedMarkdownContainer.ChildElements.ToArray();
                    renderedMarkdownContainer.RemoveAllChildren();
                    OpenXmlElement lastInsertedElement = paragraphBeforeMd;
                    for (int idx = 0; idx < convertedMdElements.Length; idx++)
                    {
                        var currentElement = convertedMdElements[idx];
                        if (idx == 0 && currentElement is Paragraph firstMdParagraph)
                        {
                            var children = firstMdParagraph.ChildElements.Where(x => x is Table or Run).ToList();
                            var targetParagraph = target.GetFirstAncestor<Paragraph>();
                            MergeStyle(targetParagraph, firstMdParagraph);
                            firstMdParagraph.RemoveAllChildren();
                            target.GetFirstAncestor<Run>().InsertAfterSelf(children);
                        }
                        else
                        {
                            lastInsertedElement = lastInsertedElement.InsertAfterSelf(currentElement);
                        }

                        if (idx == convertedMdElements.Length - 1) // last element merge with paragraph after MD
                        {
                            if (paragraphBeforeMd != paragraphAfterMd) // the original paragraph was split
                            {
                                if (lastInsertedElement is Paragraph lastInsertedParagraph) // merge runs into first paragraph
                                {
                                    var children = paragraphAfterMd.ChildElements.Where(x => x is Table or Run).ToList();
                                    paragraphAfterMd.RemoveAllChildren();
                                    paragraphAfterMd.Remove();
                                    foreach (var c in children)
                                    {
                                        lastInsertedParagraph.AppendChild(c);
                                    }
                                }
                            }
                        }
                    }
                }
                target.RemoveWithEmptyParent();
            }
            finally
            {
                m_nestingDepth--;
            }
        }

        private static void MergeStyle(Paragraph existing, Paragraph addedStyle)
        {
            if (existing == null || addedStyle == null)
            {
                return;
            }
            var existingParaProps = existing.GetFirstChild<ParagraphProperties>();
            var addedProperties = (ParagraphProperties)addedStyle.GetFirstChild<ParagraphProperties>();
            if (addedProperties != null)
            {
                if (existingParaProps != null)
                {
                    existingParaProps.ParagraphStyleId = (ParagraphStyleId)addedProperties.ParagraphStyleId?.CloneNode(true);
                    existingParaProps.ParagraphBorders = (ParagraphBorders)addedProperties.ParagraphBorders?.CloneNode(true);
                }
                else
                {
                    existing.AddChild(addedProperties.CloneNode(true));
                }
            }
        }

        private void DoVariableReplacementInParagraphs(Body mdContainer)
        {
            if (!mdContainer.InnerText.Contains('{'))
            {
                return;
            }
            var processor = new XmlNodeTemplate(mdContainer, m_processSettings, m_modelLookup, m_variableReplacer, m_scriptCompiler, m_mainDocumentPart);
            processor.Process();
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
