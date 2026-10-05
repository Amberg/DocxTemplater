using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// https://github.com/Amberg/DocxTemplater/issues/143
    /// Property names have always been resolved case-insensitively - the model prefixes
    /// passed to BindModel and the loop variables now behave the same way.
    /// </summary>
    internal class Issue143CaseInsensitiveModelPrefixTest
    {
        private static string Render(string template, string prefix, object model)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(template)))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel(prefix, model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body.InnerText;
        }

        [TestCase("{{customerDetails.Name}}")]
        [TestCase("{{CustomerDetails.Name}}")]
        [TestCase("{{CUSTOMERDETAILS.name}}")]
        public void PlaceholderResolvesPrefixIndependentOfCase(string template)
        {
            Assert.That(Render(template, "CustomerDetails", new { Name = "John" }), Is.EqualTo("John"));
        }

        [TestCase("{?{customerDetails.Flag}}yes{{/}}")]
        [TestCase("{?{CUSTOMERDETAILS.Flag}}yes{{/}}")]
        public void ExpressionResolvesPrefixIndependentOfCase(string template)
        {
            Assert.That(Render(template, "CustomerDetails", new { Flag = true }), Is.EqualTo("yes"));
        }

        [TestCase("{{#ds.Items}}{{Items.Name}}{{/}}")]
        [TestCase("{{#ds.Items}}{{items.Name}}{{/}}")]
        [TestCase("{{#DS.Items}}{{ITEMS.Name}}{{/}}")]
        public void LoopVariableResolvesIndependentOfCase(string template)
        {
            var model = new { Items = new[] { new { Name = "John" }, new { Name = "Alice" } } };
            Assert.That(Render(template, "ds", model), Is.EqualTo("JohnAlice"));
        }

        [Test]
        public void DictionaryModelResolvesIndependentOfCase()
        {
            var model = new Dictionary<string, object> { { "Name", "John" } };
            Assert.That(Render("{{ds.name}}", "ds", model), Is.EqualTo("John"));
        }
    }
}
