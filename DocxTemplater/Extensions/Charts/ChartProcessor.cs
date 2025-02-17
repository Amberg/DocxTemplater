using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Index = DocumentFormat.OpenXml.Drawing.Charts.Index;
using Path = System.IO.Path;

namespace DocxTemplater.Extensions.Charts
{
    public class ChartProcessor : ITemplateProcessorExtension
    {
        private readonly Dictionary<string, ChartReference> m_insertedChartReferences = new();

        public void PreProcess(OpenXmlCompositeElement content)
        {
            m_insertedChartReferences.Clear();
        }

        public void ReplaceVariables(ITemplateProcessingContext templateContext, OpenXmlElement parentNode, List<OpenXmlElement> newContent)
        {
            var root = parentNode.GetRoot();
            if (root is OpenXmlPartRootElement rootElement &&
                rootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
            {
                // Search for chart references in the content.
                var charts = newContent.SelectMany(x => x.Descendants<ChartReference>());
                foreach (var chartReference in charts)
                {
                    if (!m_insertedChartReferences.TryAdd(chartReference.Id, chartReference))
                    {
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
                                        ReplaceChartData(chartPart, chartData);
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
            else
            {
                throw new InvalidOperationException("ChartReference must be a child of a MainDocumentPart.");
            }
        }

        private static void ReplaceChartData(ChartPart chartPart, ChartData chartData)
        {
            var chartTitle = chartPart.ChartSpace.Descendants<Title>().FirstOrDefault();
            var title = chartTitle.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
            if (title != null)
            {
                title.Text = chartData.ChartTitle;
            }

            var barChart = chartPart.ChartSpace.Descendants<BarChart>().SingleOrDefault();
            if (barChart != null)
            {
                HandleBarChart(barChart, chartData);
            }
        }

        private static void HandleBarChart(BarChart barChart, ChartData chartData)
        {
            // replace chart series
            var series = barChart.ChildElements.OfType<BarChartSeries>().ToList();
            if (series.Count == 0)
            {
                return;
            }

            var firstSeriesEntry = (BarChartSeries)series.First().CloneNode(true);
            // rmove old series
            foreach (var old in series)
            {
                old.Remove();
            }

            BarChartSeries appendAfterThisSeris = null;
            for (int i = 0; i < chartData.Series.Count; i++)
            {
                var data = chartData.Series[i];
                var cloneBarChartSeries = CloneBarChartSeries(i, firstSeriesEntry);

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
                barChart.ValidateOpenXmlElement();
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