
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Extensions.Charts;

namespace DocxTemplater.Test
{
    internal class ChartTest
    {
        [Test]
        public void RenderTemplateWithBarChart_ModelMissing_HighlightError()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream,
                new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
            var charts = new[]
            {
                new
                {
                    Text = "Test 1"
                },
                new
                {
                    Text = "Test 2",
                }
            };
            var model = new
            {
                Items = charts
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText.Contains("Property 'MyChart' not found"));
        }

        [Test]
        public void RenderTemplateWithBarChart_ModelMissing_RemoveContent()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent });
            var charts = new[]
            {
                new
                {
                    Text = "Test 1"
                },
                new
                {
                    Text = "Test 2",
                }
            };
            var model = new
            {
                Items = charts
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();

            result.SaveAsFileAndOpenInWord();
            result.Position = 0;


            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<Drawing>().Count(), Is.EqualTo(0));
        }



        [Test]
        public void RenderTemplateWithBarChart()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var charts = new[]
            {
                new
                {
                    Text = "One serie",
                    MyChart = new ChartData()
                    {
                        ChartTitle = "Foo 1",
                        Categories = ["2022", "2023", "2024", "2025"],
                        Series =
                        [
                            new()
                            {
                                Name = "serie 1",
                                Values = [22.0, 55.0, 46.25, 90.56]
                            }
                        ]

                    }
                },
                new
                {
                    Text = "Not all cats in ervery S",
                    MyChart = new ChartData()
                    {
                        ChartTitle = "Foo 2",
                        Categories = ["Cat1", "Cat2", "Cat3", "Cat4", "Cat5"],
                        Series =
                        [
                            new()
                            {
                                Name = "serie 1",
                                Values = [2200.0, 5500.0, 4600.25, 9560.56],
                            },

                            new()
                            {
                                Name = "serie 2",
                                Values = [1200.0, 2500.0, 8600.25, 4560.56],
                            },

                            new()
                            {
                                Name = "serie 3",
                                Values = [1000.0, 2000.0, 8000.25],
                            }
                        ]
                    }
                },
                new
                {
                    Text = "Move values than categories",
                    MyChart = new ChartData()
                    {
                        ChartTitle = "Foo 2",
                        Categories = ["Cat1", "Cat2", "Cat3"],
                        Series =
                        [
                            new()
                            {
                                Name = "serie 1",
                                Values = [2200.0, 5500.0, 4600.25, 9560.56],
                            },

                            new()
                            {
                                Name = "serie 2",
                                Values = [1200.0, 2500.0, 8600.25, 4560.56],
                            },

                            new()
                            {
                                Name = "serie 3",
                                Values = [1000.0, 2000.0, 8000.25],
                            }
                        ]
                    }
                }
            };
            var model = new
            {
                Items = charts
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            // docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<ChartReference>().Count(), Is.EqualTo(3));
        }


        [Test]
        public void RenderDifferentChartType()
        {
            using var fileStream = File.OpenRead("Resources/ChartTypesTest.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.ThrowException });
            var model = new
            {
                Chart = new ChartData()
                {
                    ChartTitle = "Foo 1",
                    Categories = ["2022", "2023", "2024", "2025"],
                    Series =
                    [
                        new()
                        {
                            Name = "serie 1",
                            Values = [22.0]
                        },
                        new()
                        {
                            Name = "serie 2",
                            Values = [55.0]
                        },
                        new()
                        {
                            Name = "serie 3",
                            Values = [66.0]
                        },
                        new()
                        {
                            Name = "serie 4",
                            Values = [44]
                        },
                        new()
                        {
                            Name = "serie 5",
                            Values = [12]
                        },
                        new()
                        {
                            Name = "serie 6",
                            Values = [90]
                        }
                    ]
                }
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();

            result.SaveAsFileAndOpenInWord();
            result.Position = 0;


            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<ChartReference>().Count(), Is.EqualTo(4));
        }



        [Test]
        public void RenderChartWithoutSeries()
        {
            using var fileStream = File.OpenRead("Resources/ChartTypesTest.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.ThrowException });
            var model = new
            {
                Chart = new ChartData()
                {
                    ChartTitle = "Foo 1",
                    Categories = ["2022", "2023", "2024", "2025"],
                    Series = null,
                }
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();

            result.SaveAsFileAndOpenInWord();
            result.Position = 0;


            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<ChartReference>().Count(), Is.EqualTo(4));
        }

        [Test]
        public void RenderChartWithoutCategories()
        {
            using var fileStream = File.OpenRead("Resources/ChartTypesTest.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.ThrowException });
            var model = new
            {
                Chart = new ChartData()
                {
                    ChartTitle = "Foo 1",
                    Categories = null,
                    Series =
                    [
                        new(){
                            Name = "serie",
                            Values = [1, 2, 3]
                        },
                    ]
                }
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();

            result.SaveAsFileAndOpenInWord();
            result.Position = 0;


            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<ChartReference>().Count(), Is.EqualTo(4));
        }

        // Regression guard for the data-binding itself (not just that the chart parts exist).
        // ChartTypesTest.docx contains 4 bar/bar3D charts bound to ds.Chart via the spreadsheet path.
        // We assert the bound series names, categories and values actually land in every chart part.
        [Test]
        public void RenderBarChart_BindsSeriesNamesCategoriesAndValues()
        {
            using var fileStream = File.OpenRead("Resources/ChartTypesTest.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.ThrowException });
            var model = new
            {
                Chart = new ChartData()
                {
                    ChartTitle = "Foo",
                    Categories = ["Alpha", "Beta", "Gamma"],
                    Series =
                    [
                        new() { Name = "Series A", Values = [10.0, 20.0, 30.0] },
                        new() { Name = "Series B", Values = [40.0, 50.0, 60.0] }
                    ]
                }
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            result.Position = 0;

            using var document = WordprocessingDocument.Open(result, false);
            var mainPart = document.MainDocumentPart;
            var chartParts = mainPart.Document.Body.Descendants<ChartReference>()
                .Select(r => (ChartPart)mainPart.GetPartById(r.Id))
                .ToList();
            Assert.That(chartParts, Has.Count.EqualTo(4));

            string[] expectedCategories = ["Alpha", "Beta", "Gamma"];
            double[] expectedValuesA = [10.0, 20.0, 30.0];
            double[] expectedValuesB = [40.0, 50.0, 60.0];

            foreach (var chartPart in chartParts)
            {
                var series = chartPart.ChartSpace.Descendants<BarChartSeries>().ToList();
                Assert.That(series, Has.Count.EqualTo(2), "expected one BarChartSeries per bound series");

                // no leftover cat/val from the cloned template series
                Assert.That(series[0].Elements<CategoryAxisData>().Count(), Is.EqualTo(1));
                Assert.That(series[0].Elements<Values>().Count(), Is.EqualTo(1));

                Assert.That(GetSeriesName(series[0]), Is.EqualTo("Series A"));
                Assert.That(GetSeriesName(series[1]), Is.EqualTo("Series B"));

                Assert.That(GetCategoryTexts(series[0]), Is.EqualTo(expectedCategories));
                Assert.That(GetCategoryTexts(series[1]), Is.EqualTo(expectedCategories));

                Assert.That(GetValueNumbers(series[0]), Is.EqualTo(expectedValuesA));
                Assert.That(GetValueNumbers(series[1]), Is.EqualTo(expectedValuesB));
            }
        }

        // Pie/Pie3D show only the first series; Doughnut shows all of them.
        [TestCase("Resources/PieChart.docx", 1)]
        [TestCase("Resources/PieChart3d.docx", 1)]
        [TestCase("Resources/Doughnut.docx", 3)]
        public void RenderPieFamilyChart_BindsSeriesAndVariesColors(string templatePath, int expectedSeriesCount)
        {
            using var fileStream = File.OpenRead(templatePath);
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.ThrowException });
            var model = new
            {
                Items = new[]
                {
                    new
                    {
                        Text = "Pie",
                        MyChart = new ChartData()
                        {
                            ChartTitle = "Pie Title",
                            Categories = ["A", "B", "C", "D"],
                            Series =
                            [
                                new() { Name = "Series A", Values = [10.0, 20.0, 30.0, 40.0] },
                                new() { Name = "Series B", Values = [11.0, 21.0, 31.0, 41.0] },
                                new() { Name = "Series C", Values = [12.0, 22.0, 32.0, 42.0] }
                            ]
                        }
                    }
                }
            };

            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;

            string[] expectedCategories = ["A", "B", "C", "D"];
            double[] expectedValuesA = [10.0, 20.0, 30.0, 40.0];

            using var document = WordprocessingDocument.Open(result, false);
            var mainPart = document.MainDocumentPart;
            var chartPart = mainPart.Document.Body.Descendants<ChartReference>()
                .Select(r => (ChartPart)mainPart.GetPartById(r.Id))
                .Single();

            var series = chartPart.ChartSpace.Descendants<PieChartSeries>().ToList();
            Assert.That(series, Has.Count.EqualTo(expectedSeriesCount), "pie shows first series only; doughnut shows all");

            // first (visible) series content
            Assert.That(GetSeriesName(series[0]), Is.EqualTo("Series A"));
            Assert.That(series[0].Elements<CategoryAxisData>().Count(), Is.EqualTo(1));
            Assert.That(series[0].Elements<Values>().Count(), Is.EqualTo(1));
            Assert.That(GetCategoryTexts(series[0]), Is.EqualTo(expectedCategories));
            Assert.That(GetValueNumbers(series[0]), Is.EqualTo(expectedValuesA));

            // stale per-slice colors are removed; the chart varies colors automatically
            Assert.That(series.SelectMany(s => s.Elements<DataPoint>()).Count(), Is.EqualTo(0), "template dPt entries should be stripped");
            var chartElement = chartPart.ChartSpace.Descendants<OpenXmlCompositeElement>()
                .First(e => e is PieChart or Pie3DChart or DoughnutChart);
            Assert.That(chartElement.GetFirstChild<VaryColors>()?.Val?.Value, Is.True, "varyColors must be enabled");

            // series must keep their schema position: varyColors before ser, ser before holeSize.
            var children = chartElement.ChildElements.ToList();
            int firstSeriesIndex = children.FindIndex(c => c is PieChartSeries);
            int varyColorsIndex = children.FindIndex(c => c is VaryColors);
            Assert.That(varyColorsIndex, Is.LessThan(firstSeriesIndex), "varyColors must precede the series");
            var holeSize = chartElement.GetFirstChild<HoleSize>();
            if (holeSize != null)
            {
                int lastSeriesIndex = children.FindLastIndex(c => c is PieChartSeries);
                Assert.That(lastSeriesIndex, Is.LessThan(children.IndexOf(holeSize)), "series must precede holeSize");
            }
        }

        private static string GetSeriesName(OpenXmlCompositeElement series)
        {
            return series.GetFirstChild<SeriesText>()?.Descendants<StringPoint>().FirstOrDefault()?.NumericValue?.Text
                   ?? series.GetFirstChild<SeriesText>()?.Descendants<NumericValue>().FirstOrDefault()?.Text;
        }

        // The current code appends fresh cat/val onto a clone that already carries the template's
        // cat/val, so use the last element - it is the generated one both before and after the refactor.
        private static string[] GetCategoryTexts(OpenXmlCompositeElement series)
        {
            var cat = series.Elements<CategoryAxisData>().LastOrDefault();
            return cat?.Descendants<StringPoint>().Select(p => p.NumericValue?.Text).ToArray() ?? [];
        }

        private static double[] GetValueNumbers(OpenXmlCompositeElement series)
        {
            var val = series.Elements<Values>().LastOrDefault();
            return val?.Descendants<NumericPoint>()
                .Select(p => double.Parse(p.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                .ToArray() ?? [];
        }
    }
}
