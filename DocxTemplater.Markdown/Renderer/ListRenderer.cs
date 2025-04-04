using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ListRenderer : OpenXmlObjectRenderer<ListBlock>
    {
        private int m_level = -1;
        private int m_levelWithSameOrdering = -1;
        private bool? m_lastLevelOrdered;
        private readonly ListStyleFactory m_orderedListStyleFactory;
        private readonly ListStyleFactory m_unorderedListStyleFactory;

        public ListRenderer(MainDocumentPart mainDocumentPart, MarkDownFormatterConfiguration configuration)
        {
            m_orderedListStyleFactory = new ListStyleFactory(true, configuration, mainDocumentPart);
            m_unorderedListStyleFactory = new ListStyleFactory(false, configuration, mainDocumentPart);
        }

        protected override void Write(MarkdownToOpenXmlRenderer renderer, ListBlock listBlock)
        {

            var listStyleFactory = listBlock.IsOrdered ? m_orderedListStyleFactory : m_unorderedListStyleFactory;
            listStyleFactory.EnsureExists();

            StartListLevel(listBlock.IsOrdered);
            listStyleFactory.EnsureLevelDefinitionExists(m_levelWithSameOrdering);
            try
            {
                if (m_level == 0)
                {
                    renderer.ReplaceIfCurrentParagraphIsEmpty(new Paragraph());
                }

                foreach (var item in listBlock)
                {

                    var numberingProps =
                        new NumberingProperties(
                            new NumberingLevelReference() { Val = m_levelWithSameOrdering },
                            new NumberingId() { Val = listStyleFactory.Numbering.NumberID }
                        );
                    var listParagraph = new Paragraph();
                    var paragraphProperties = new ParagraphProperties(numberingProps)
                    {
                        ParagraphStyleId = new ParagraphStyleId()
                        {
                            Val = listStyleFactory.ListParagraphStyle
                        }
                    };
                    listParagraph.ParagraphProperties = paragraphProperties;
                    renderer.ReplaceIfCurrentParagraphIsEmpty(listParagraph);
                    var listItem = (ListItemBlock)item;
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


        private void StartListLevel(bool ordered)
        {
            m_level++;
            if (m_lastLevelOrdered != ordered)
            {
                m_levelWithSameOrdering = -1;
            }
            m_levelWithSameOrdering++;
            m_lastLevelOrdered = ordered;
        }


        private void EndListLevel()
        {
            m_level--;
            m_levelWithSameOrdering--;
        }
    }
}
