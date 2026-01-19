using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Parsers;
using System;
using System.Linq;
using Markdig.Extensions.Tables;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace DocxTemplater.Markdown
{
    public class MarkdownFormatter : IFormatter
    {
        private readonly MarkDownFormatterConfiguration m_configuration;
        private int m_nestingDepth;
        private static readonly MarkdownPipeline MarkdownPipeline;

        public MarkdownFormatter(MarkDownFormatterConfiguration configuration = null)
        {
            m_configuration = configuration ?? MarkDownFormatterConfiguration.Default;
        }

        static MarkdownFormatter()
        {
            var builder = new MarkdownPipelineBuilder();
            builder.UseGridTables()
                .UsePipeTables(new PipeTableOptions { InferColumnWidthsFromSeparator = true })
                .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
                .UseFigures();
            MarkdownPipeline = builder.Build();
        }

        public bool CanHandle(Type type, string prefix)
        {
            string prefixUpper = prefix.ToUpper();
            return prefixUpper is "MD" && type == typeof(string);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Value is not string mdText)
            {
                return;
            }

            if (m_nestingDepth > 3)
            {
                throw new OpenXmlTemplateException("Markdown nesting depth exceeded");
            }

            var contextSpecificConfiguration = m_configuration.Clone();
            if (formatterContext.Args.Length > 0)
            {
                var arguments = HelperFunctions.ParseArguments(formatterContext.Args);
                if (arguments.TryGetValue("ts", out var tableStyleName))
                {
                    contextSpecificConfiguration.TableStyle = tableStyleName;
                }
                if (arguments.TryGetValue("ls", out var listStyle))
                {
                    contextSpecificConfiguration.OrderedListStyle = listStyle;
                }
                if (arguments.TryGetValue("ols", out var orderedListStyle))
                {
                    contextSpecificConfiguration.OrderedListStyle = orderedListStyle;
                }
            }

            m_nestingDepth++;
            try
            {
                var root = target.GetRoot();
                if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
                {
                    // markdown treat multiple newlines as a single newline, we insert a space to keep the newlines
                    mdText = mdText.Replace("\r\n", "\r");
                    var markdownDocument = MarkdownParser.Parse(mdText, MarkdownPipeline);

                    // split the paragraph at the target
                    var split = target.GetFirstAncestor<Paragraph>().SplitAfterElement(target);
                    Paragraph paragraphBeforeMd = split.OfType<Paragraph>().First();
                    Paragraph paragraphAfterMd = split.OfType<Paragraph>().Last();

                    // create a container for the markdown as render target
                    var renderedMarkdownContainer = new Body();
                    var containerParagraph = new Paragraph();
                    renderedMarkdownContainer.Append(containerParagraph);

                    var renderer = new MarkdownToOpenXmlRenderer(containerParagraph, target, templateContext.MainDocumentPart, contextSpecificConfiguration, templateContext.ImageService);
                    renderer.Render(markdownDocument);

#if DEBUG
                    Console.WriteLine(renderer.MarkdownStructureAsString);
#endif

                    try
                    {
                        DoVariableReplacementInParagraphs(renderedMarkdownContainer, templateContext);
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
                                    var children = paragraphAfterMd.ChildElements.Where(x => x is Table or Run
                                        || InsertionPoint.HasAlreadyInsertionPointMarker(x)
                                        || x.IsMarked()
                                        ).ToList();
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
            var addedProperties = addedStyle.GetFirstChild<ParagraphProperties>();
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

        private static void DoVariableReplacementInParagraphs(Body mdContainer, ITemplateProcessingContext templateContext)
        {
            if (!mdContainer.InnerText.Contains('{'))
            {
                return;
            }
            var processor = new XmlNodeTemplate(mdContainer, (ITemplateProcessingContextAccess)templateContext);
            processor.Process();
        }
    }
}
