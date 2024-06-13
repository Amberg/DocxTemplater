using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;


namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ListStyleFactory
    {
        private readonly MarkDownFormatterConfiguration m_markDownFormatterConfiguration;
        private readonly MainDocumentPart m_mainDocumentPart;
        private AbstractNum m_abstractNum;
        private NumberingDefinitionsPart m_numberingDefinitionsPart;

        public ListStyleFactory(bool ordered, MarkDownFormatterConfiguration markDownFormatterConfiguration, MainDocumentPart mainDocumentPart)
        {
            m_markDownFormatterConfiguration = markDownFormatterConfiguration;
            m_mainDocumentPart = mainDocumentPart;
            Ordered = ordered;
        }

        public bool Ordered { get; }

        public NumberingInstance Numbering { get; private set; }

        public string ListParagraphStyle
        {
            get;
            private set;
        }

        public void EnsureExists()
        {
            if (Numbering != null)
            {
                return;
            }

            m_numberingDefinitionsPart = m_mainDocumentPart.NumberingDefinitionsPart;
            // Ensure the main document part has a NumberingDefinitionsPart
            if (m_numberingDefinitionsPart == null)
            {
                m_numberingDefinitionsPart = m_mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                m_numberingDefinitionsPart.Numbering = new Numbering();
            }

            var part = m_mainDocumentPart.StyleDefinitionsPart;
            if (part == null)
            {
                part = m_mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                part.Styles = new Styles();
            }

            CrateListParagraphStyleIfNotExistent();
            if (!TryFindNumberingStyle())
            {
                CrateNumberingInstance();
            }
        }

        public void EnsureLevelDefinitionExists(int level)
        {
            var levelElement = m_abstractNum.ChildElements.OfType<Level>().FirstOrDefault(x => x.LevelIndex == level);
            if (levelElement == null)
            {
                var levelConfig = Ordered ? m_markDownFormatterConfiguration.OrderedListLevelConfiguration : m_markDownFormatterConfiguration.UnorderedListLevelConfiguration;
                var config = levelConfig[level % levelConfig.Count];
                levelElement = new Level()
                {
                    LevelIndex = level,
                    NumberingFormat = new NumberingFormat()
                    {
                        Val = config.NumberingFormat
                    },
                    LevelText = new LevelText() { Val = config.LevelText },
                    StartNumberingValue = new StartNumberingValue() { Val = 1 },
                };
                var paraProps = new PreviousParagraphProperties(new Indentation()
                {
                    Left = (config.IndentPerLevel * (level + 1)).ToString(),
                    Hanging = level % 2 == 0 ? "360" : "180"
                });
                levelElement.Append(paraProps);
                if (!string.IsNullOrEmpty(config.FontOverride))
                {
                    levelElement.Append(new NumberingSymbolRunProperties(new RunFonts() { Ascii = config.FontOverride, HighAnsi = config.FontOverride }));

                }
                m_abstractNum.Append(levelElement);
                m_mainDocumentPart.NumberingDefinitionsPart.Numbering.Save();
            }
        }

        private bool TryFindNumberingStyle()
        {
            var styleName = Ordered ? m_markDownFormatterConfiguration.OrderedListStyle : m_markDownFormatterConfiguration.UnorderedListStyle;
            var numberingStyle = m_mainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().FirstOrDefault(x => x.StyleName?.Val == styleName);
            if (numberingStyle == null)
            {
                return false;
            }
            var styleId = numberingStyle.StyleId;
            // find abstract numberings with this style
            m_abstractNum = m_mainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<AbstractNum>()
                .Where(r => r.StyleLink != null && r.StyleLink.Val == styleId).MaxBy(x => x.ChildElements.Count);
            if (m_abstractNum == null)
            {
                return false;
            }
            Numbering = m_mainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<NumberingInstance>().FirstOrDefault(x => x.AbstractNumId?.Val == m_abstractNum.AbstractNumberId) ??
                        m_mainDocumentPart.NumberingDefinitionsPart.Numbering.CreateNewNumberingInstance(m_abstractNum.AbstractNumberId);
            m_numberingDefinitionsPart.Numbering.Save();
            return Numbering != null && m_abstractNum != null;
        }


        /// <summary>
        /// Crates a AbstractNum format for multiple levels of lists
        /// </summary>
        /// <returns></returns>
        private void CrateNumberingInstance()
        {
            m_abstractNum = m_numberingDefinitionsPart.Numbering.CreateNewAbstractNumbering();
            Numbering = m_numberingDefinitionsPart.Numbering.CreateNewNumberingInstance(m_abstractNum.AbstractNumberId);
            m_numberingDefinitionsPart.Numbering.Save();
        }

        private void CrateListParagraphStyleIfNotExistent()
        {
            if (ListParagraphStyle == null)
            {
                var stylePart = m_mainDocumentPart.StyleDefinitionsPart;
                var style = stylePart.Styles?.Elements<Style>().FirstOrDefault(x => x.StyleId == "ListParagraph");
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
                    stylePart.Styles?.Append(style);
                    stylePart.Styles?.Save();
                }
                ListParagraphStyle = "ListParagraph";
            }
        }
    }
}
