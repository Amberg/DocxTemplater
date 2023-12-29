using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Templates;
using System.Diagnostics;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Path = System.IO.Path;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXml.Teplates.Test
{
    internal class DocxTemplateTest
    {
        [Test]
        public void ReplaceTextBoldIsPreserved()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text("This Value:")),
                new Run(new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }), new Text("{{Property1}}")),
                new Run(new Text("will be replaced"))
            )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.AddModel("Property1", "Replaced");
            var result = docTemplate.Process();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            // check that bold is preserved
            Assert.That(body.Descendants<Bold>().First().Val, Is.EqualTo(OnOffValue.FromBoolean(true)));
            // check that text is replaced
            Assert.That(body.Descendants<Text>().Skip(1).First().Text, Is.EqualTo("Replaced"));
        }

        [Test]
        public void BindCollection()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("This Value:")),
                    new Run(new Text("{{#ds.Items}}")),
                    new Run(new Text("For each run {{ds.Items.Name}}"))
                ),
            new Paragraph(
                    new Run(new Text("{{ds.Items.Value}}")),
                    new Run(new Text("{{#ds.Items.InnerCollection}}")),
                        new Run(new Text("{{ds.Items.Value}}")),
                        new Run(new Text("{{ds.Items.InnerCollection.Name}}")),
                        new Run(new Text("{{ds.Items.InnerCollection.InnerValue}}")),
                    new Run(new Text("{{/ds.Items.InnerCollection}}")),
                    new Run(new Text("{{/ds.Items}}")),
                    new Run(new Text("will be replaced"))
                )
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.AddModel("ds",
                new {
                        PropertInRoot = "RootValue", 
                        Items = new[]
                        {
                            new {Name = "Item1", Value = "Value1", InnerCollection = new[] {new {Name = "InnerItem1", InnerValue = "InnerValue2"}}},
                            new {Name = "Item2", Value = "Value2", InnerCollection = new[] {new {Name = "InnerItem3", InnerValue = "InnerValue4"}}},
                        }
                });
            var result = docTemplate.Process();
            Assert.IsNotNull(result);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            // check values have been replaced
            Assert.That(body.InnerText, Is.EqualTo("This Value:For each run Item1Value1Value1InnerItem1InnerValue2Value1InnerItem1InnerValue2For each run Item1Value1Value1InnerItem1InnerValue2Value1InnerItem1InnerValue2will be replaced"));

            //SaveAsFileAndOpen(result);
        }


        private void SaveAsFileAndOpen(Stream stream)
        {
            var fileName = Path.ChangeExtension(Path.GetTempFileName(), "docx");
            using (var fileStream = File.OpenWrite(fileName))
            {
                stream.CopyTo(fileStream);
            }

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = fileName;
            psi.UseShellExecute = true;
            using var proc = Process.Start(psi);
            proc.WaitForExit();
        }
    }


}
