using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Model;

namespace DocxTemplater.Test
{
    class ExpressionParserWithCustom_ITemplateModelTest
    {
        [Test]
        public void ExpressionWithCustomTemplateModel()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new Run(new Text("{{#ds.Items}}" +
                                 "{?{.CustomModel.MyBoolProp}}There{{/}}" +
                                 "{{/ds.Items}}"))
            )));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new
            {
                Items = new[]
                {
                    new { CustomModel = new TestModel() }
                }
            });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("There"));
        }

        private class TestModel : ITemplateModel
        {
            public bool TryGetPropertyValue(string propertyName, out ValueWithMetadata value)
            {
                value = new ValueWithMetadata(true);
                return true;
            }
        }
    }
}
