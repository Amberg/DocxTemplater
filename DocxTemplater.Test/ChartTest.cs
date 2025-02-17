
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Extensions.Charts;

namespace DocxTemplater.Test
{
    internal class ChartTest
    {
        [Test]
        public void RenderTemplateWithBarChart_ModelMissingHighlightError()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { BindingErrorHandling = BindingErrorHandling.HighlightErrorsInDocument });
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
        public void RenderTemplateWithBarChart_RemoveContent()
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
        public void RenderTemplateWithBarChart_WithBoundData()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var charts = new[]
            {
                new
                {
                    Text = "Test 1",
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
                    Text = "Test 2",
                    MyChart = new ChartData()
                    {
                        ChartTitle = "Foo 2",
                        Categories = ["Cat1", "Cat2", "Cat3", "Cat4", "Cat5"],
                        Series =
                        [
                            new()
                            {
                                Name = "serie 2",
                                Values = [2200.0, 5500.0, 4600.25, 9560.56],
                            },

                            new()
                            {
                                Name = "serie 3",
                                Values = [1200.0, 2500.0, 8600.25, 4560.56],
                            },

                            new()
                            {
                                Name = "serie 4",
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
            Assert.That(body.Descendants<ChartReference>().Count(), Is.EqualTo(2));
        }
    }
}
