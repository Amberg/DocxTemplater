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

        // Describes how a concrete chart type is bound. The series child elements (cat/val/tx/...)
        // are shared OpenXml types across chart kinds; only the chart element, the series element
        // type, the number of usable series and the coloring differ.
        private sealed class ChartTypeDescriptor
        {
            public ChartTypeDescriptor(Type chartElementType, Type seriesType, int maxSeries,
                Action<OpenXmlCompositeElement, OpenXmlCompositeElement, int, int> applyColoring)
            {
                ChartElementType = chartElementType;
                SeriesType = seriesType;
                MaxSeries = maxSeries;
                ApplyColoring = applyColoring;
            }

            public Type ChartElementType { get; }

            public Type SeriesType { get; }

            /// <summary>
            /// Maximum number of series the chart can display (1 for pie/pie3d, unlimited otherwise).
            /// </summary>
            public int MaxSeries { get; }

            /// <summary>
            /// Applies type-specific coloring. Receives the chart element, the series, the series
            /// index and the number of data points (slices) in the series.
            /// </summary>
            public Action<OpenXmlCompositeElement, OpenXmlCompositeElement, int, int> ApplyColoring { get; }
        }

        // Registry of supported chart types. Add an entry to support a new chart kind.
        private static readonly ChartTypeDescriptor[] s_supportedChartTypes =
        [
            new(typeof(BarChart), typeof(BarChartSeries), int.MaxValue, ColorBarSeries),
            new(typeof(Bar3DChart), typeof(BarChartSeries), int.MaxValue, ColorBarSeries),
            new(typeof(PieChart), typeof(PieChartSeries), 1, ColorPieSeries),
            new(typeof(Pie3DChart), typeof(PieChartSeries), 1, ColorPieSeries),
            new(typeof(DoughnutChart), typeof(PieChartSeries), int.MaxValue, ColorPieSeries),
        ];
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
                    if (!m_insertedChartReferences.TryAdd(chartReference.Id, chartReference))
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
                            var found = templateContext.ProcessSettings.PatternMatcher.FindSyntaxPatterns(title.Text).FirstOrDefault();
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

                                        var plotArea = chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>()?.PlotArea;
                                        var (chartElement, descriptor) = FindChart(plotArea);
                                        if (chartElement == null)
                                        {
                                            continue; // unsupported chart type - leave untouched
                                        }

                                        var externalData = chartPart.ChartSpace.GetFirstChild<ExternalData>();
                                        if (externalData?.Id != null) //  add check it is spead- sheet
                                        {
                                            var part = chartPart.GetPartById(externalData.Id);
                                            if (part is EmbeddedPackagePart embeddedPart && embeddedPart.ContentType.Equals(spreadsheetContentType, StringComparison.OrdinalIgnoreCase))
                                            {
                                                string sheetName = SpreadSheetHelper.ReplaceDataInSpreadSheet(embeddedPart, chartData);
                                                PopulateSeriesWithReferences(chartElement, descriptor, chartData, sheetName);
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            HandleLiteralChart(chartElement, descriptor, chartData);
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

        // Locates the first supported chart element under the given root and returns it together
        // with its descriptor. Returns (null, null) when no supported chart type is found.
        private static (OpenXmlCompositeElement chartElement, ChartTypeDescriptor descriptor) FindChart(OpenXmlCompositeElement root)
        {
            if (root == null)
            {
                return (null, null);
            }
            foreach (var element in root.Descendants<OpenXmlCompositeElement>())
            {
                var descriptor = Array.Find(s_supportedChartTypes, d => d.ChartElementType == element.GetType());
                if (descriptor != null)
                {
                    return (element, descriptor);
                }
            }
            return (null, null);
        }

        // The spreadsheet path: the chart data lives in an embedded workbook and the series only
        // reference its cells (plus a cache for display). Works for any chart type via the descriptor.
        private static void PopulateSeriesWithReferences(OpenXmlCompositeElement chartElement, ChartTypeDescriptor descriptor, ChartData chartData, string sheetName)
        {
            int seriesCount = Math.Min(chartData.Series.Count, descriptor.MaxSeries);
            var clonedSeries = CloneSeriesAndRemoveOld(chartElement, descriptor, seriesCount, out var anchor);
            if (clonedSeries == null)
            {
                return;
            }
            for (int i = 0; i < seriesCount; i++)
            {
                var series = clonedSeries[i];
                var data = chartData.Series[i];
                string col = SpreadSheetHelper.GetColumnName(i + 1);  // 1 → A, 2 → B, 3 → C ...

                // index, order and series name cell ${col}1 - replaces any cloned template data
                ClearSeriesData(series);
                SetSeriesHeader(series, (uint)i, new SeriesText(
                    new StringReference(
                        new Formula($"{sheetName}!${col}$1"),
                        new StringCache(
                            new PointCount { Val = 1 },
                            new StringPoint { Index = 0, NumericValue = new NumericValue(data.Name) }
                        )
                    )
                ));

                // categories from A2:A{n+1}
                var catCount = (uint)chartData.Categories.Count();
                var strCache = new StringCache();
                strCache.Append(new PointCount { Val = catCount });
                for (uint idx = 0; idx < catCount; idx++)
                {
                    strCache.Append(new StringPoint { Index = idx, NumericValue = new NumericValue(chartData.Categories.ElementAt((int)idx)) });
                }

                series.Append(new CategoryAxisData(new StringReference(new Formula($"{sheetName}!$A$2:$A${catCount + 1}"), strCache)));

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

                descriptor.ApplyColoring(chartElement, series, i, (int)valCount);
                anchor = InsertSeriesAfter(chartElement, series, anchor);
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


        // Clones the template series of the descriptor's series type (cloning preserves the concrete
        // CLR type) and removes the originals. Returns null when the template has no series to clone.
        // <paramref name="insertAnchor"/> receives the sibling the new series must be inserted after
        // (the element preceding the original series, or null if the series were the first child) so
        // the regenerated series land in their schema-correct position instead of at the very end.
        private static OpenXmlCompositeElement[] CloneSeriesAndRemoveOld(OpenXmlCompositeElement chartElement, ChartTypeDescriptor descriptor, int seriesCount, out OpenXmlElement insertAnchor)
        {
            var series = chartElement.ChildElements.Where(e => e.GetType() == descriptor.SeriesType).Cast<OpenXmlCompositeElement>().ToList();
            if (series.Count == 0)
            {
                insertAnchor = null;
                return null;
            }
            insertAnchor = series[0].PreviousSibling();
            var clones = new OpenXmlCompositeElement[seriesCount];
            for (int i = 0; i < seriesCount; i++)
            {
                var template = i < series.Count ? series[i] : series[^1];
                clones[i] = (OpenXmlCompositeElement)template.CloneNode(true);
            }
            // remove old series
            foreach (var old in series)
            {
                old.Remove();
            }
            return clones;
        }

        // Inserts a regenerated series after the given anchor (preserving schema order) and returns
        // the inserted series so it can serve as the anchor for the next one.
        private static OpenXmlElement InsertSeriesAfter(OpenXmlCompositeElement chartElement, OpenXmlElement series, OpenXmlElement anchor)
        {
            if (anchor == null)
            {
                chartElement.InsertAt(series, 0);
            }
            else
            {
                chartElement.InsertAfter(series, anchor);
            }
            return series;
        }

        // The literal path: no embedded workbook, so categories and values are embedded directly
        // into the chart XML as string/number literals. Works for any chart type via the descriptor.
        private static void HandleLiteralChart(OpenXmlCompositeElement chartElement, ChartTypeDescriptor descriptor, ChartData chartData)
        {
            int seriesCount = Math.Min(chartData.Series.Count, descriptor.MaxSeries);
            var clonedSeries = CloneSeriesAndRemoveOld(chartElement, descriptor, seriesCount, out var anchor);
            if (clonedSeries == null)
            {
                return;
            }
            for (int i = 0; i < seriesCount; i++)
            {
                var data = chartData.Series[i];
                var series = clonedSeries[i];

                // preserve the template's number format before clearing the cloned data
                var formatCode = (FormatCode)series.Descendants<FormatCode>().FirstOrDefault()?.CloneNode(true) ?? new FormatCode("General");

                ClearSeriesData(series);
                SetSeriesHeader(series, (uint)i, new SeriesText(new NumericValue(data.Name)));

                var stringLiteral = new StringLiteral()
                {
                    PointCount = new PointCount() { Val = new UInt32Value((uint)data.Values.Count()) },
                };
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
                        NumericValue = new NumericValue(dataItem.ToString("F1", CultureInfo.InvariantCulture)),
                        Index = new UInt32Value((uint)counter)
                    });
                    stringLiteral.AppendChild(new StringPoint()
                    {
                        NumericValue = new NumericValue(GetCategoryName(chartData, counter)),
                        Index = new UInt32Value((uint)counter)
                    });
                    counter++;
                }

                series.Append(new CategoryAxisData() { StringLiteral = stringLiteral });
                series.Append(new Values() { NumberLiteral = numberLiteral });
                descriptor.ApplyColoring(chartElement, series, i, data.Values.Count());

                anchor = InsertSeriesAfter(chartElement, series, anchor);
            }
        }

        // Removes the index/order/series-name and re-adds them at the front in schema order, then
        // sets fresh values. idx/order/tx are the first three children of every chart series type.
        private static void SetSeriesHeader(OpenXmlCompositeElement series, uint index, SeriesText seriesText)
        {
            series.GetFirstChild<Index>()?.Remove();
            series.GetFirstChild<Order>()?.Remove();
            series.GetFirstChild<SeriesText>()?.Remove();
            series.InsertAt(seriesText, 0);
            series.InsertAt(new Order { Val = index }, 0);
            series.InsertAt(new Index { Val = index }, 0);
        }

        // Removes any category/value data carried over from the cloned template series.
        private static void ClearSeriesData(OpenXmlCompositeElement series)
        {
            foreach (var cat in series.Elements<CategoryAxisData>().ToList())
            {
                cat.Remove();
            }
            foreach (var val in series.Elements<Values>().ToList())
            {
                val.Remove();
            }
        }

        // Bar/Bar3D: each series gets a distinct accent color cycling through the theme.
        private static void ColorBarSeries(OpenXmlCompositeElement chartElement, OpenXmlCompositeElement series, int index, int pointCount)
        {
            var solidFill = series.GetFirstChild<ChartShapeProperties>()?.GetFirstChild<SolidFill>();
            if (solidFill?.SchemeColor != null)
            {
                solidFill.SchemeColor.Val = (index % 6) switch
                {
                    0 => SchemeColorValues.Accent1,
                    1 => SchemeColorValues.Accent2,
                    2 => SchemeColorValues.Accent3,
                    3 => SchemeColorValues.Accent4,
                    4 => SchemeColorValues.Accent5,
                    5 => SchemeColorValues.Accent6,
                    _ => SchemeColorValues.Accent1
                };
            }
        }

        // Pie/Pie3D/Doughnut: slices are colored per data point (dPt). The template author stays in
        // control of those colors, so the cloned dPt are kept - only entries referring to slices that
        // no longer exist (idx >= number of slices) are dropped. varyColors lets Word auto-color any
        // slices the template did not define a color for.
        private static void ColorPieSeries(OpenXmlCompositeElement chartElement, OpenXmlCompositeElement series, int index, int pointCount)
        {
            foreach (var dataPoint in series.Elements<DataPoint>().ToList())
            {
                if (dataPoint.Index?.Val != null && dataPoint.Index.Val.Value >= (uint)pointCount)
                {
                    dataPoint.Remove();
                }
            }

            var varyColors = chartElement.GetFirstChild<VaryColors>();
            if (varyColors == null)
            {
                chartElement.InsertAt(new VaryColors() { Val = true }, 0); // varyColors is the chart element's first child
            }
            else
            {
                varyColors.Val = true;
            }
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