using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using ExternalData = DocumentFormat.OpenXml.Drawing.Charts.ExternalData;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using NumericValue = DocumentFormat.OpenXml.Drawing.Charts.NumericValue;
using Path = System.IO.Path;
using Values = DocumentFormat.OpenXml.Drawing.Charts.Values;

namespace DocxTemplater.Extensions.Charts
{
    public class ChartProcessor : ITemplateProcessorExtension
    {
        private readonly Dictionary<string, ChartReference> m_insertedChartReferences = new();
        private const string spreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public void PreProcess(OpenXmlCompositeElement content)
        {
            m_insertedChartReferences.Clear();
        }

        public void ReplaceVariables(ITemplateProcessingContext templateContext, OpenXmlElement parentNode, List<OpenXmlElement> newContent)
        {
            var root = parentNode.GetRoot();
            if (root is OpenXmlPartRootElement rootElement && rootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
            {
                // Search for chart references in the content.
                var charts = newContent.SelectMany(x => x.Descendants<ChartReference>());
                foreach (var chartReference in charts)
                {
                    if (!m_insertedChartReferences.ContainsKey(chartReference.Id))
                    {
                        m_insertedChartReferences.Add(chartReference.Id, chartReference);
                    }
                    else
                    {
                        // if this chart already used - clone it
                        CloneChart(chartReference, mainDocumentPart);
                    }

                    // get the chart
                    var chartPart = (ChartPart)mainDocumentPart.GetPartById(chartReference.Id);

                    // find the chart title
                    var chartTitle = chartPart.ChartSpace.Descendants<Title>().FirstOrDefault();
                    if (chartTitle != null)
                    {
                        // replace the chart title
                        var title = chartTitle.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                        if (title != null)
                        {
                            var found = PatternMatcher.FindSyntaxPatterns(title.Text).FirstOrDefault();
                            if (found != null && found.Type == PatternType.Variable)
                            {
                                try
                                {

                                    var model = templateContext.ModelLookup.GetValue(found.Variable);
                                    if (model is ChartData chartData)
                                    {
                                        chartData.Series ??= [new ChartSeries() { Name = "-", Values = [] }]; // no data fallback
                                        chartData.Categories ??= []; // no data fallback

                                        ReplaceChartTitle(chartPart, chartData);

                                        var externalData = chartPart.ChartSpace.GetFirstChild<ExternalData>();
                                        if (externalData?.Id != null) //  add check it is spead- sheet
                                        {
                                            var part = chartPart.GetPartById(externalData.Id);
                                            if (part is EmbeddedPackagePart embeddedPart && embeddedPart.ContentType.Equals(spreadsheetContentType, StringComparison.OrdinalIgnoreCase))
                                            {
                                                string sheetName = SpreadSheetHelper.ReplaceDataInSpreadSheet(embeddedPart, chartData);
                                                PopulateSeriesWithReferences(chartPart, chartData, sheetName);
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            HandleBarChart(GetBarChartChild(chartPart.ChartSpace), chartData);
                                        }
                                    }
                                }
                                catch (OpenXmlTemplateException e)
                                {
                                    if (templateContext.ProcessSettings.BindingErrorHandling == BindingErrorHandling.SkipBindingAndRemoveContent)
                                    {
                                        chartReference.GetFirstAncestor<Drawing>()?.Remove();
                                    }
                                    else if (templateContext.ProcessSettings.BindingErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                                    {
                                        templateContext.VariableReplacer.AddError(e.Message);
                                    }
                                    else
                                    {
                                        throw;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static OpenXmlCompositeElement GetBarChartChild(OpenXmlCompositeElement root)
        {
            if (root == null)
            {
                return null;
            }
            return (OpenXmlCompositeElement)root.Descendants<BarChart>().SingleOrDefault() ?? root.Descendants<Bar3DChart>().SingleOrDefault();
        }

        private static void PopulateSeriesWithReferences(ChartPart chartPart, ChartData chartData, string sheetName)
        {
            var barChart = GetBarChartChild(chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>()?.PlotArea);
            if (barChart == null)
            {
                return;
            }
            var clonedBarChartSeriesCopies = CloneBarChartSeriesAndRemoveOld(barChart, chartData);
            for (uint i = 0; i < chartData.Series.Count; i++)
            {

                var series = clonedBarChartSeriesCopies[i];
                var data = chartData.Series[(int)i];
                string col = SpreadSheetHelper.GetColumnName((int)i + 1);  // 1 → A, 2 → B, 3 → C ...

                // index and order
                series.Index = new Index { Val = i };
                series.Order = new Order { Val = i };

                // serie name cell ${col}1
                series.SeriesText = new SeriesText(
                    new StringReference(
                        new Formula($"{sheetName}!${col}$1"),
                        new StringCache(
                            new PointCount { Val = 1 },
                            new StringPoint { Index = 0, NumericValue = new NumericValue(data.Name) }
                        )
                    )
                );

                // categories from A2:A{n+1}
                var catCount = (uint)chartData.Categories.Count();
                var strCache = new StringCache();
                strCache.Append(new PointCount { Val = catCount });
                for (uint idx = 0; idx < catCount; idx++)
                {
                    strCache.Append(new StringPoint { Index = idx, NumericValue = new NumericValue(chartData.Categories.ElementAt((int)idx)) });
                }

                var axisData = new CategoryAxisData(new StringReference(new Formula($"{sheetName}!$A$2:$A${catCount + 1}"), strCache));
                series.Append(axisData);

                // values aus {col}2:{col}{n+1}
                var valCount = (uint)data.Values.Count();
                var numCache = new NumberingCache();
                numCache.Append(new FormatCode("General"));
                numCache.Append(new PointCount { Val = valCount });
                for (uint idx = 0; idx < valCount; idx++)
                {
                    numCache.Append(new NumericPoint { Index = idx, NumericValue = new NumericValue(data.Values.ElementAt((int)idx).ToString("F1", CultureInfo.InvariantCulture)) });
                }

                series.Append(new Values(
                    new NumberReference(
                        new Formula($"{sheetName}!${col}$2:${col}${valCount + 1}"),
                        numCache
                    )
                ));

                barChart.Append(series);

            }
        }

        private static void ReplaceChartTitle(ChartPart chartPart, ChartData chartData)
        {
            var chartTitle = chartPart.ChartSpace.Descendants<Title>().FirstOrDefault();
            var title = chartTitle.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
            if (title != null)
            {
                title.Text = chartData.ChartTitle;
            }
        }


        private static BarChartSeries[] CloneBarChartSeriesAndRemoveOld(OpenXmlCompositeElement barChart, ChartData chartData)
        {
            // replace chart series
            var series = barChart.ChildElements.OfType<BarChartSeries>().ToList();
            BarChartSeries[] clonedBarChartSeriesCopies = new BarChartSeries[chartData.Series.Count];
            for (int i = 0; i < clonedBarChartSeriesCopies.Length; i++)
            {
                if (i < series.Count)
                {
                    clonedBarChartSeriesCopies[i] = CloneBarChartSeries(i, (BarChartSeries)series.ElementAt(i));
                }
                else
                {
                    clonedBarChartSeriesCopies[i] = CloneBarChartSeries(i, series.Last());
                }
            }
            // rmove old series
            foreach (var old in series)
            {
                old.Remove();
            }
            return clonedBarChartSeriesCopies;
        }

        private static void HandleBarChart(OpenXmlCompositeElement barChart, ChartData chartData)
        {
            var clonedBarChartSeriesCopies = CloneBarChartSeriesAndRemoveOld(barChart, chartData);
            BarChartSeries appendAfterThisSeris = null;
            for (int i = 0; i < chartData.Series.Count; i++)
            {
                var data = chartData.Series[i];
                var cloneBarChartSeries = clonedBarChartSeriesCopies[i];

                cloneBarChartSeries.Index = new Index() { Val = new UInt32Value((uint)i) };
                cloneBarChartSeries.Order = new Order() { Val = new UInt32Value((uint)i) };
                cloneBarChartSeries.SeriesText = new SeriesText(new NumericValue(data.Name));
                var stringLiteral = new StringLiteral()
                {
                    PointCount = new PointCount() { Val = new UInt32Value((uint)data.Values.Count()) },
                };
                var catData = new CategoryAxisData()
                {
                    StringLiteral = stringLiteral
                };
                cloneBarChartSeries.AddChild(catData);

                var copiedValues = cloneBarChartSeries.GetFirstChild<Values>();
                var formatCode = copiedValues?.GetFirstChild<FormatCode>() ?? new FormatCode("General");
                var numberLiteral = new NumberLiteral
                {
                    FormatCode = formatCode,
                    PointCount = new PointCount() { Val = new UInt32Value((uint)data.Values.Count()) }
                };
                int counter = 0;
                foreach (var dataItem in data.Values)
                {
                    numberLiteral.AppendChild(new NumericPoint()
                    {
                        NumericValue = new NumericValue(dataItem.ToString("F1")),
                        Index = new UInt32Value((uint)counter)
                    });
                    stringLiteral.AppendChild(new StringPoint()
                    {
                        NumericValue = new NumericValue(GetCategoryName(chartData, counter)),
                        Index = new UInt32Value((uint)counter)
                    });
                    counter++;
                }

                copiedValues.AddChild(numberLiteral);
                if (appendAfterThisSeris == null)
                {
                    barChart.AddChild(cloneBarChartSeries);
                    appendAfterThisSeris = cloneBarChartSeries;
                }
                else
                {
                    barChart.InsertAfter(cloneBarChartSeries, appendAfterThisSeris);
                    appendAfterThisSeris = cloneBarChartSeries;
                }
#if DEBUG
                // charts are valid and can be opened in word but the validation fails
                // barChart.ValidateOpenXmlElement();
#endif
            }
        }


        private static BarChartSeries CloneBarChartSeries(int i, BarChartSeries firstSeriesEntry)
        {
            var result = (BarChartSeries)firstSeriesEntry.CloneNode(true);
            // search properties

            var solidFill = result.ChartShapeProperties?.GetFirstChild<SolidFill>();
            if (solidFill?.SchemeColor != null)
            {
                var value = (i % 6) switch
                {
                    0 => SchemeColorValues.Accent1,
                    1 => SchemeColorValues.Accent2,
                    2 => SchemeColorValues.Accent3,
                    3 => SchemeColorValues.Accent4,
                    4 => SchemeColorValues.Accent5,
                    5 => SchemeColorValues.Accent6,
                    _ => SchemeColorValues.Accent1
                };
                solidFill.SchemeColor.Val = value;
            }

            return result;
        }


        private static string GetCategoryName(ChartData data, int index)
        {
            if (data.Categories == null || !data.Categories.Any())
            {
                return string.Empty;
            }
            if (data.Categories.Count() > index)
            {
                return data.Categories.ElementAt(index);
            }

            return data.Categories.Last();
        }

        private void CloneChart(ChartReference chartReference, MainDocumentPart mainDocumentPart)
        {
            // Get the OpenXmlPartRootElement from the chart reference.

            string newId = $"{chartReference.Id}_{m_insertedChartReferences.Count}";
            // Get the original ChartPart using the relationship ID.
            var originalChartPart = (ChartPart)mainDocumentPart.GetPartById(chartReference.Id);

            // Create a new ChartPart (this will generate a new target path).
            var clonedChartPart = mainDocumentPart.AddNewPart<ChartPart>(newId);

            // Copy the main content of the chart.
            using (var sourceStream = originalChartPart.GetStream(FileMode.Open))
            using (var targetStream = clonedChartPart.GetStream(FileMode.Create))
            {
                sourceStream.CopyTo(targetStream);
            }

            // Clone the relationships from the original ChartPart.
            CopyPartRelationshipsRecursive(originalChartPart, clonedChartPart, mainDocumentPart);

            // Get the new relationship ID and update the chart reference.
            var newRelId = mainDocumentPart.GetIdOfPart(clonedChartPart);
            chartReference.Id = newRelId;

            m_insertedChartReferences.Add(newId, chartReference);
        }

        private static void CopyPartRelationshipsRecursive(OpenXmlPart source, OpenXmlPart target,
            MainDocumentPart mainDocumentPart)
        {
            foreach (var rel in source.Parts)
            {
                OpenXmlPart clonedPart = null;
                if (rel.OpenXmlPart is EmbeddedPackagePart)
                {
                    var uri = rel.OpenXmlPart.Uri;
                    // get extension
                    var extension = Path.GetExtension(uri.OriginalString);
                    clonedPart =
                        mainDocumentPart.AddEmbeddedPackagePart(
                            new PartTypeInfo(rel.OpenXmlPart.ContentType, extension), rel.RelationshipId);
                    var id = mainDocumentPart.GetIdOfPart(clonedPart);
                    target.AddPart(clonedPart, id);

                }
                else if (rel.OpenXmlPart is ChartColorStylePart)
                {
                    clonedPart = target.AddNewPart<ChartColorStylePart>();

                }
                else if (rel.OpenXmlPart is ChartStylePart)
                {
                    clonedPart = target.AddNewPart<ChartStylePart>();
                }

                // copy content
                using (var sourceStream = rel.OpenXmlPart.GetStream(FileMode.Open))
                using (var targetStream = new MemoryStream())
                {
                    sourceStream.CopyTo(targetStream);
                    targetStream.Position = 0; // Reset stream position before feeding data.
                    clonedPart.FeedData(targetStream);
                }

                CopyPartRelationshipsRecursive(rel.OpenXmlPart, clonedPart, mainDocumentPart);
            }
        }
    }
}