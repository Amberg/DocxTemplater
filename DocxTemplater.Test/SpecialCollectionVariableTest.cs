using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocxTemplater.Test
{
    internal class SpecialCollectionVariableTest
    {
        [Test]
        public void TestIndexVariableInLoop()
        {
            var model = new[] { "Item1", "Item2", "Item3", "Item4" };
            var template = "Items:{{#Items}}{{Items._Idx}}{{.}} {{/Items}}";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(template)))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("Items", model);
            var result = docTemplate.Process();
            result.Position = 0;
            // compare body
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                                  "<w:r>" +
                                                  "<w:t xml:space=\"preserve\">Items:</w:t><w:t xml:space=\"preserve\">1</w:t><w:t xml:space=\"preserve\">Item1</w:t>" +
                                                  "<w:t xml:space=\"preserve\"> </w:t><w:t xml:space=\"preserve\">2</w:t><w:t xml:space=\"preserve\">Item2</w:t>" +
                                                  "<w:t xml:space=\"preserve\"> </w:t><w:t xml:space=\"preserve\">3</w:t><w:t xml:space=\"preserve\">Item3</w:t>" +
                                                  "<w:t xml:space=\"preserve\"> </w:t><w:t xml:space=\"preserve\">4</w:t><w:t xml:space=\"preserve\">Item4</w:t>" +
                                                  "<w:t xml:space=\"preserve\"> </w:t></w:r></w:p>"));
        }

        [Test]
        public void TestConditionWithIndexVariableInLoop()
        {
            var model = new[] { "Item1", "Item2", "Item3", "Item4" };
            var template = "Items:{{#Items}}{?{Items._Idx % 2 == 0}}{{.}}{{/}}{{/Items}}";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(template)))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("Items", model);
            var result = docTemplate.Process();
            result.Position = 0;
            // compare body
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t xml:space=\"preserve\">Items:" +
                                                  "</w:t><w:t xml:space=\"preserve\">Item2</w:t>" +
                                                  "<w:t xml:space=\"preserve\">Item4</w:t></w:r></w:p>"));
        }
    }
}
