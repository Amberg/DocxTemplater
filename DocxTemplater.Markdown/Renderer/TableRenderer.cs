using System.Linq;
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

            m_tableStyle ??= FindDefaultTableStyle(m_mainDocument, m_markDownFormatterConfiguration);

            if (m_tableStyle != null)
            {
                tableProperties.TableStyle = new WP.TableStyle
                {
                    Val = m_tableStyle.StyleId
                };
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

        public static WP.Style FindDefaultTableStyle(MainDocumentPart mainDocumentPart, MarkDownFormatterConfiguration markDownFormatterConfiguration)
        {
            var part = mainDocumentPart.StyleDefinitionsPart;
            if (part?.Styles == null)
            {
                return null;
            }

            // First search for style by name
            if (markDownFormatterConfiguration.TableStyle != null)
            {
                var style = mainDocumentPart.FindTableStyleByName(markDownFormatterConfiguration.TableStyle);
                if (style != null)
                {
                    return style;
                }
            }

            // 1. Fallback: Use the style from an existing table in the document.
            var firstTable = mainDocumentPart.Document?.Body?.Elements<WP.Table>().FirstOrDefault();
            if (firstTable != null)
            {
                var tblPr = firstTable.GetFirstChild<WP.TableProperties>();
                var tblStyle = tblPr?.GetFirstChild<WP.TableStyle>();
                if (tblStyle != null && !string.IsNullOrEmpty(tblStyle.Val?.Value))
                {
                    var existingStyle = part.Styles.Elements<WP.Style>()
                        .FirstOrDefault(s => s.StyleId == tblStyle.Val.Value);
                    if (existingStyle != null)
                    {
                        return existingStyle;
                    }
                }
            }

            // 3. Fallback: Use the latent default style ("TableGrid").
            var latentDefault = part.Styles.Elements<WP.Style>().LastOrDefault(s => s.Type == WP.StyleValues.Table);
            if (latentDefault != null)
            {
                return latentDefault;
            }

            // 4. As a last resort, return any table style marked as default.
            return part.Styles.Elements<WP.Style>().FirstOrDefault(s => s.Type == WP.StyleValues.Table && s.Default != null && s.Default.Value);
        }
    }
}
