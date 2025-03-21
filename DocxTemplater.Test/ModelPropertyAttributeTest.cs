﻿using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Model;

namespace DocxTemplater.Test
{
    internal class ModelPropertyAttributeTest
    {

        [Test]
        public void FormatterFromAttribute()
        {
            var content = "Hello {{ds.Name}} {{ds.LastName}} - {{ds.LastName}:ToUpper} - {{ds.Number}}";
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(content)))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream, new ProcessSettings { Culture = new CultureInfo("en-US") });
            docTemplate.BindModel("ds", new TestModel()
            {
                Name = "John",
                LastName = "Doe",
                Number = 1m / 3m
            });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            // validate content
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerText, Is.EqualTo("Hello JOHN doe - DOE - 0.33"));
        }

        private class TestModel
        {
            [ModelProperty(DefaultFormatter = "toupper")]
            public string Name { get; set; }

            [ModelProperty(DefaultFormatter = "tolower")]
            public string LastName { get; set; }

            [ModelProperty(DefaultFormatter = "F(n2)")]
            public decimal Number { get; set; }
        }
    }
}
