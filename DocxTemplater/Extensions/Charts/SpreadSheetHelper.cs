using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace DocxTemplater.Extensions.Charts
{
    internal class SpreadSheetHelper
    {
        public static string ReplaceDataInSpreadSheet(EmbeddedPackagePart embeddedPart, ChartData chartData)
        {
            string sheetName = "Sheet1";

            using var memStream = new MemoryStream();
            using (var spreadsheet = SpreadsheetDocument.Create(memStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());


                var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);


                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // first row: series names
                var headerRow = new Row();
                headerRow.AppendChild(CreateTextCell("A1", ""));
                for (int i = 0; i < chartData.Series.Count; i++)
                {
                    headerRow.AppendChild(CreateTextCell(GetCellReference(i + 1, 1), chartData.Series[i].Name));
                }
                sheetData.AppendChild(headerRow);

                // categories and values
                var categories = chartData.Categories.ToList();
                for (int rowIndex = 0; rowIndex < categories.Count; rowIndex++)
                {
                    var row = new Row();
                    row.AppendChild(CreateTextCell(GetCellReference(0, rowIndex + 2), categories[rowIndex]));

                    for (int colIndex = 0; colIndex < chartData.Series.Count; colIndex++)
                    {
                        var series = chartData.Series[colIndex];
                        double value = series.Values.ElementAtOrDefault(rowIndex);
                        row.AppendChild(CreateNumberCell(GetCellReference(colIndex + 1, rowIndex + 2), value));
                    }

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }


            memStream.Position = 0;
            embeddedPart.FeedData(memStream);

            return sheetName;
        }

        private static Cell CreateTextCell(string cellReference, string cellValue)
        {
            return new Cell()
            {
                CellReference = cellReference,
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };
        }

        private static Cell CreateNumberCell(string cellReference, double cellValue)
        {
            return new Cell()
            {
                CellReference = cellReference,
                DataType = CellValues.Number,
                CellValue = new CellValue(cellValue.ToString(CultureInfo.InvariantCulture))
            };
        }

        private static string GetCellReference(int columnIndex, int rowIndex)
        {
            return GetColumnName(columnIndex) + rowIndex;
        }

        /*
         * Helper function to convert a column index (0-based) to an Excel-style column name (A, B, C, ..., Z, AA, AB, ...)
         */
        public static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
    }
}
