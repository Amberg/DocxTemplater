using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class DynamicTableBlock : ContentBlock
    {
        private readonly string m_tablenName;

        public DynamicTableBlock(string tablenName, VariableReplacer variableReplacer)
            : base(variableReplacer)
        {
            m_tablenName = tablenName;
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            var model = models.GetValue(m_tablenName);
            if (model is IDynamicTable dynamicTable)
            {

                var headersName = $"{m_tablenName}.{nameof(IDynamicTable.Headers)}";
                var columnsName = $"{m_tablenName}.{nameof(IDynamicTable.Rows)}";

                var table = m_content.OfType<Table>().FirstOrDefault();
                var headerRow = table?.Elements<TableRow>().FirstOrDefault(row => row.Descendants<Text>().Any(d => d.HasMarker(PatternType.Variable) && d.Text.Contains($"{{{{{headersName}")));
                var headerCell = headerRow?.Elements<TableCell>().FirstOrDefault();

                var dataRow = table?.Elements<TableRow>().FirstOrDefault(row => row.Descendants<Text>().Any(d => d.HasMarker(PatternType.Variable) && d.Text.Contains($"{{{{{columnsName}")));
                var dataCell = dataRow?.Elements<TableCell>().FirstOrDefault(row => row.Descendants<Text>().Any(d => d.HasMarker(PatternType.Variable) && d.Text.Contains($"{{{{{columnsName}")));
                if (headerCell == null || dataCell == null)
                {
                    throw new OpenXmlTemplateException($"Dynamic table block must contain exactly one table with at least a header and a data row");
                }

                // write headers
                foreach (var header in dynamicTable.Headers.Reverse())
                {
                    using var headerScope = models.OpenScope();
                    headerScope.AddVariable(headersName, header);
                    var clonedCell = headerCell.CloneNode(true);
                    headerCell.InsertAfterSelf(clonedCell);
                    m_variableReplacer.ReplaceVariables(clonedCell);
                    ExpandChildBlocks(models, parentNode);
                }
                // remove header cell
                headerCell.Remove();

                // write data
                var lastRow = dataRow;
                var cellInsertionPoint = InsertionPoint.CreateForElement(dataCell);
                foreach (var row in dynamicTable.Rows)
                {
                    TableRow clonedRow = (TableRow)dataRow.CloneNode(true);
                    lastRow.InsertAfterSelf(clonedRow);
                    lastRow = clonedRow;

                    var insertion = cellInsertionPoint.GetElement(clonedRow);
                    foreach (var column in row.Reverse())
                    {
                        using var columnScope = models.OpenScope();
                        columnScope.AddVariable(columnsName, column);
                        var clonedCell = dataCell.CloneNode(true);
                        insertion.InsertAfterSelf(clonedCell);
                        m_variableReplacer.ReplaceVariables(clonedCell);
                        ExpandChildBlocks(models, parentNode);
                    }
                    insertion.Remove();
                }
                dataRow.Remove();
                dataCell.Remove();

                // ensure all rows have the same number of cells
                var maxCells = dynamicTable.Rows.Max(r => r.Count());
                foreach (var row in table.Elements<TableRow>())
                {
                    var cells = row.Elements<TableCell>().ToList();
                    while (cells.Count < maxCells)
                    {
                        var cell = (TableCell)cells.Last().CloneNode(true);
                        cells.Last().InsertAfterSelf(cell);
                        cells.Add(cell);
                    }
                }

                InsertContent(parentNode, new List<OpenXmlElement> { table });
            }
            else
            {
                throw new OpenXmlTemplateException($"Value of {m_tablenName} is not of type {typeof(IDynamicTable)}");
            }
        }

        public override void SetContent(InsertionPoint insertionPoint, IReadOnlyCollection<OpenXmlElement> blockContent)
        {
            var tables = blockContent.OfType<Table>().ToList();
            if (tables.Count != 1)
            {
                throw new OpenXmlTemplateException($"Dynamic table block must contain exactly one table, but found {tables.Count}");
            }
            base.SetContent(insertionPoint, tables);
        }

        public override string ToString()
        {
            return $"Dynamic Table: {m_tablenName}";
        }
    }
}
