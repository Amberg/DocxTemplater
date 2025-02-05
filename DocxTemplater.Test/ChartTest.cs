using DocxTemplater.Charts;
using System.IO;

namespace DocxTemplater.Test
{
    internal class ChartTest
    {
        [Test]
        public void RenderTemplateWithBarChart()
        {
            using var fileStream = File.OpenRead("Resources/BarChart.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterExtension(new ChartProcessor());
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
            docTemplate.Process();

            var result = docTemplate.Process();
            var fileName = "C:\\Work\\DocxTemplater\\DocxTemplater.Test\\Resources\\output.docx";
            File.Delete(fileName);
            using (var outStream = File.OpenWrite(fileName))
            {
                result.CopyTo(outStream);
            }
        }
    }
}
