using DocumentFormat.OpenXml.Packaging;
using Markdig.Extensions.Tables;
using Table = Markdig.Extensions.Tables.Table;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class TableRenderer : OpenXmlObjectRenderer<Table>
    {
        private readonly MarkDownFormatterConfiguration m_markDownFormatterConfiguration;
        private readonly MainDocumentPart m_mainDocument;
        private WP.Style m_tableStyle;

        public TableRenderer(MarkDownFormatterConfiguration markDownFormatterConfiguration, MainDocumentPart mainDocument)
        {
            m_markDownFormatterConfiguration = markDownFormatterConfiguration;
            m_mainDocument = mainDocument;
        }

        private static string PercentToFiftiethsOfAPercent(double percent)
        {
            // 5000 fiftieths-of-a-percent = 100% - strange openxml units
            return ((int)(percent * 50)).ToString();
        }

        protected override void Write(MarkdownToOpenXmlRenderer renderer, Table mkTable)
        {
            var table = new WP.Table();
            var tableProperties = new WP.TableProperties
            {
                TableWidth = new WP.TableWidth
                {
                    Type = WP.TableWidthUnitValues.Pct,
                    Width = PercentToFiftiethsOfAPercent(100)
                }
            };

            if (m_markDownFormatterConfiguration.TableStyle != null)
            {
                if (m_tableStyle == null)
                {
                    m_tableStyle ??= m_mainDocument.FindTableStyleByName(m_markDownFormatterConfiguration.TableStyle);
                    tableProperties.TableStyle = new WP.TableStyle { Val = m_markDownFormatterConfiguration.TableStyle };
                }
            }

            var tableGrid = new WP.TableGrid();
            foreach (var _ in mkTable.ColumnDefinitions)
            {
                /* TODO:
                // Add full support for alignment as defined in specs of Pipe table
                // https://github.com/xoofx/markdig/blob/master/src/Markdig.Tests/Specs/PipeTableSpecs.md
                */
                tableGrid.Append(new WP.GridColumn());
            }

            table.AppendChild(tableProperties);
            table.Append(tableGrid);

            foreach (TableRow row in mkTable)
            {
                var tableRow = new WP.TableRow();
                table.AppendChild(tableRow);
                for (int i = 0; i < row.Count; i++)
                {
                    var cell = (TableCell)row[i];
                    var cellProperties = new WP.TableCellProperties();
                    var cellWidth = new WP.TableCellWidth { Type = WP.TableWidthUnitValues.Auto };
                    cellProperties.Append(cellWidth);
                    var cellElement = new WP.TableCell(cellProperties);
                    tableRow.AppendChild(cellElement);

                    // cell paragraph
                    var cellParagraph = new WP.Paragraph();
                    cellElement.Append(cellParagraph);
                    using var paragraphScope = renderer.PushParagraph(cellParagraph);
                    renderer.WriteChildren(cell);
                }
            }
            renderer.AddParagraph(table);
            renderer.ExplicitParagraph = true;
            renderer.AddParagraph();
        }
    }
}
