﻿using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class DynamicTableBlock : ContentBlock
    {
        private readonly string m_tablenName;

        public DynamicTableBlock(IVariableReplacer variableReplacer, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(variableReplacer, patternType, startTextNode, startMatch)
        {
            m_tablenName = startMatch.Variable;
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            var model = models.GetValue(m_tablenName);
            if (model is IDynamicTable dynamicTable)
            {
                if (!dynamicTable.Headers.Any())
                {
                    return;
                }

                var headersName = $"{m_tablenName}.{nameof(IDynamicTable.Headers)}";
                var columnsName = $"{m_tablenName}.{nameof(IDynamicTable.Rows)}";

                // kind of a hack to get the table from the child block
                // TODO: refactor this to create wrapper block as ContentBlock and DynamicTableBlock as child
                var child = m_childBlocks.Single();
                var childContent = m_childBlocks.Single().Content;
                var table = childContent.OfType<Table>().FirstOrDefault();
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
                    child.ExpandChildBlocks(models, parentNode);
                }
                // remove header cell
                headerCell.Remove();

                // write data
                var lastRow = dataRow;
                var cellInsertionPoint = InsertionPoint.CreateForElement(dataCell, "dc");
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
                        child.ExpandChildBlocks(models, parentNode);
                    }
                    insertion.Remove();
                }
                dataRow.Remove();
                dataCell.Remove();

                // ensure all rows have the same number of cells
                var maxCells = dynamicTable.Rows.DefaultIfEmpty().Max(r => r?.Count() ?? 0);
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
                throw new OpenXmlTemplateException($"'{m_tablenName}' is not of type {typeof(IDynamicTable)}");
            }
        }

        public override void Validate()
        {
            base.Validate();
            if (m_childBlocks.Count != 1)
            {
                throw new OpenXmlTemplateException($"Dynamic table block must contain exactly one child block");
            }
        }

        public override string ToString()
        {
            return $"Dynamic Table: {m_tablenName}";
        }
    }
}
