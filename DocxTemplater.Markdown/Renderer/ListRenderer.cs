using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using System.Linq;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ListRenderer : OpenXmlObjectRenderer<ListBlock>
    {
        private int m_level = -1;
        private int m_levelWithSameOrdering = -1;
        private bool? m_lastLevelOrdered;

        private readonly MainDocumentPart m_mainDocumentPart;
        private readonly MarkDownFormatterConfiguration m_configuration;
        private AbstractNum m_currentAbstractNumNotOrdered;
        private AbstractNum m_currentAbstractNumOrdered;
        private NumberingInstance m_currentNumberingInstanceOrdered;
        private NumberingInstance m_currentNumberingInstanceNotOrdered;
        private NumberingDefinitionsPart m_numberingDefinitionsPart;
        private string m_listParagraphStyle;

        public ListRenderer(MainDocumentPart mainDocumentPart, MarkDownFormatterConfiguration configuration)
        {
            m_mainDocumentPart = mainDocumentPart;
            m_configuration = configuration;
        }

        protected override void Write(MarkdownToOpenXmlRenderer renderer, ListBlock listBlock)
        {

            StartListLevel(listBlock.IsOrdered);
            try
            {
                var numberingInstance = listBlock.IsOrdered ? m_currentNumberingInstanceOrdered : m_currentNumberingInstanceNotOrdered;

                foreach (var item in listBlock)
                {
                    var numberingProps =
                    new NumberingProperties(
                        new NumberingLevelReference() { Val = m_levelWithSameOrdering },
                        new NumberingId() { Val = numberingInstance.NumberID }
                    );
                    var listItem = (ListItemBlock)item;
                    var listParagraph = new Paragraph();
                    var paragraphProperties = new ParagraphProperties(numberingProps)
                    {
                        ParagraphStyleId = new ParagraphStyleId() { Val = m_listParagraphStyle }
                    };
                    listParagraph.ParagraphProperties = paragraphProperties;
                    renderer.ReplaceIfCurrentParagraphIsEmpty(listParagraph);
                    renderer.ExplicitParagraph = true;
                    renderer.WriteChildren(listItem);
                }
            }
            finally
            {
                EndListLevel();
            }

            if (m_level == -1)
            {
                renderer.AddParagraph();
            }
        }

        private void CrateListParagraphStyleIfNotExistent()
        {
            if (m_listParagraphStyle == null)
            {
                var part = m_mainDocumentPart.StyleDefinitionsPart;
                if (part == null)
                {
                    part = m_mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    part.Styles = new Styles();
                }

                var style = part.Styles?.Elements<Style>().FirstOrDefault(x => x.StyleId == "ListParagraph");
                if (style == null)
                {
                    style = new Style()
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = "ListParagraph",
                        CustomStyle = true,
                        StyleName = new StyleName() { Val = "List Paragraph" },
                        UIPriority = new UIPriority() { Val = 34 },
                        PrimaryStyle = new PrimaryStyle(),
                        Rsid = new Rsid() { Val = "004004FF" }
                    };
                    var styleParagraphProperties = new StyleParagraphProperties();
                    styleParagraphProperties.Append(new ContextualSpacing());
                    style.Append(styleParagraphProperties);
                    part.Styles?.Append(style);
                    part.Styles?.Save();
                }
                m_listParagraphStyle = "ListParagraph";
            }
        }

        private void StartListLevel(bool ordered)
        {
            m_level++;
            if (m_level == 0)
            {
                CrateListParagraphStyleIfNotExistent();
                CrateNumberingInstance(ordered);
            }

            if (m_lastLevelOrdered != ordered)
            {
                m_levelWithSameOrdering = -1;
            }
            m_levelWithSameOrdering++;
            m_lastLevelOrdered = ordered;

            // try to find level
            var abstractNumberDefinition = GetAbstractNumberingDefinition(ordered);
            var level = abstractNumberDefinition.ChildElements.OfType<Level>().FirstOrDefault(x => x.LevelIndex == m_level);
            if (level == null)
            {
                var levelConfig = ordered ? m_configuration.OrderedListLevelConfiguration : m_configuration.UnorderedListLevelConfiguration;
                var config = levelConfig[m_levelWithSameOrdering % levelConfig.Count];
                level = new Level()
                {
                    LevelIndex = m_level,
                    NumberingFormat = new NumberingFormat()
                    {
                        Val = config.NumberingFormat
                    },
                    LevelText = new LevelText() { Val = config.LevelText },
                    StartNumberingValue = new StartNumberingValue() { Val = 1 },
                };
                var paraProps = new PreviousParagraphProperties(new Indentation()
                {
                    Left = (config.IndentPerLevel * (m_level + 1)).ToString(),
                    Hanging = m_level % 2 == 0 ? "360" : "180"
                });
                level.Append(paraProps);
                if (!string.IsNullOrEmpty(config.FontOverride))
                {
                    level.Append(new NumberingSymbolRunProperties(new RunFonts() { Ascii = config.FontOverride, HighAnsi = config.FontOverride }));

                }
                abstractNumberDefinition.Append(level);
                m_numberingDefinitionsPart.Numbering.Save();
            }
        }

        private void EndListLevel()
        {
            m_level--;
            m_levelWithSameOrdering--;
        }

        /// <summary>
        /// Crates a AbstractNum format for multiple levels of lists
        /// </summary>
        /// <returns></returns>
        private void CrateNumberingInstance(bool ordered)
        {
            // Ensure the main document part has a NumberingDefinitionsPart
            if (m_mainDocumentPart.NumberingDefinitionsPart == null)
            {
                m_numberingDefinitionsPart = m_mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                m_numberingDefinitionsPart.Numbering = new Numbering();
            }
            else
            {
                m_numberingDefinitionsPart = m_mainDocumentPart.NumberingDefinitionsPart;
            }

            var abstractNum = GetAbstractNumberingDefinition(ordered);

            if (abstractNum == null)
            {
                abstractNum = m_numberingDefinitionsPart.Numbering.CreateNewAbstractNumbering();
                var numberIngInstance = m_numberingDefinitionsPart.Numbering.CreateNewNumberingInstance(abstractNum.AbstractNumberId);
                if (!ordered)
                {
                    m_currentAbstractNumNotOrdered = abstractNum;
                    m_currentNumberingInstanceNotOrdered = numberIngInstance;
                }
                else
                {
                    m_currentAbstractNumOrdered = abstractNum;
                    m_currentNumberingInstanceOrdered = numberIngInstance;
                }
                m_numberingDefinitionsPart.Numbering.Save();
            }

        }

        private AbstractNum GetAbstractNumberingDefinition(bool ordered)
        {
            var abstractNum = ordered ? m_currentAbstractNumOrdered : m_currentAbstractNumNotOrdered;
            return abstractNum;
        }
    }
}
