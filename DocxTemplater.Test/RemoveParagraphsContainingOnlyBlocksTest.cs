using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class RemoveParagraphsContainingOnlyBlocksTest
    {
        [Test]
        public void TestRemoveParagraphsAroundConditionalBlocks()
        {
            var paragraph1 = new Paragraph(new Run(new Text("This is first paragraph")));
            var paragraph2 = new Paragraph(new Run(new Text("{?{.Item!=null}}"))); // Condition start
            var paragraph3 = new Paragraph(new Run(new Text("Item: {{.Item.Name}}"))); // Content within condition
            var paragraph4 = new Paragraph(new Run(new Text("{{/}}"))); // Condition end
            var paragraph5 = new Paragraph(new Run(new Text("This is last paragraph")));

            var body = new Body();
            body.Append(paragraph1, paragraph2, paragraph3, paragraph4, paragraph5);

            // Process with removing paragraphs enabled - should remove empty paragraphs
            {
                using var memStream = new MemoryStream();
                using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(body.CloneNode(true));
                wpDocument.Save();
                memStream.Position = 0;

                var template = new DocxTemplate(memStream);
                template.Settings.RemoveParagraphsContainingOnlyBlocks = true;
                template.BindModel("", new { Item = (object)null });
                var result = template.Process();

                using var processedDocument = WordprocessingDocument.Open(result, false);
                // There should be 2 paragraphs (first and last) - empty one should be removed
                var paragraphs = processedDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
                Assert.That(paragraphs.Count, Is.EqualTo(2));

                // Verify the text content
                Assert.That(paragraphs[0].InnerText, Is.EqualTo("This is first paragraph"));
                Assert.That(paragraphs[1].InnerText, Is.EqualTo("This is last paragraph"));
            }
        }

        [Test]
        public void TestRemoveParagraphsAroundEmptyCollections()
        {
            var paragraph1 = new Paragraph(new Run(new Text("This is first paragraph")));
            var paragraph2 = new Paragraph(new Run(new Text("{{#Items}}"))); // Collection start
            var paragraph3 = new Paragraph(new Run(new Text("Item: {{.}}"))); // Content within collection
            var paragraph4 = new Paragraph(new Run(new Text("{{/}}"))); // Collection end
            var paragraph5 = new Paragraph(new Run(new Text("This is last paragraph")));

            var body = new Body();
            body.Append(paragraph1, paragraph2, paragraph3, paragraph4, paragraph5);

            // Process with removing paragraphs enabled - should remove empty paragraphs
            {
                using var memStream = new MemoryStream();
                using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(body.CloneNode(true));
                wpDocument.Save();
                memStream.Position = 0;

                var template = new DocxTemplate(memStream);
                template.Settings.RemoveParagraphsContainingOnlyBlocks = true;
                template.BindModel("", new { Items = Array.Empty<string>() });
                var result = template.Process();

                using var processedDocument = WordprocessingDocument.Open(result, false);
                // There should be 2 paragraphs (first and last) - empty one should be removed
                var paragraphs = processedDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
                Assert.That(paragraphs.Count, Is.EqualTo(2));

                // Verify the text content
                Assert.That(paragraphs[0].InnerText, Is.EqualTo("This is first paragraph"));
                Assert.That(paragraphs[1].InnerText, Is.EqualTo("This is last paragraph"));
            }
        }

        [Test]
        public void TestTemplateDocumentWithAllBlockTypes()
        {
            using var fileStream = File.OpenRead("Resources/RemoveParagraphsContainingOnlyBlocks.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.Settings.RemoveParagraphsContainingOnlyBlocks = true;

            // Create test data that matches the template structure
            var testData = new
            {
                Val = "Test Value",
                Items = new[] { "Item 1", "Item 2" },
                NoItems = Array.Empty<string>(),
                Models = new[]
                {
                    new { Header = "First Header", Text = "This is the first text block with some detailed content" },
                    new { Header = "Second Header", Text = "Another text block with different content" },
                    new { Header = "Third Header", Text = "Yet another block of text to test the template" },
                    new { Header = "Fourth Header", Text = "Final text block with unique content" }
                }
            };

            docTemplate.BindModel("", testData);
            var result = docTemplate.Process();
            docTemplate.Validate();

            // Save the output file for manual inspection
            var outputPath = Path.GetFullPath("RemoveParagraphsContainingOnlyBlocks_Output.docx");
            using (var fs = File.Create(outputPath))
            {
                result.CopyTo(fs);
            }
            Console.WriteLine($"Output file saved to: {outputPath}");

            // Verify the document structure
            result.Position = 0;
            using var processedDocument = WordprocessingDocument.Open(result, false);
            var paragraphs = processedDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            // The document should have removed paragraphs that only contained blocks
            // and kept paragraphs with actual content
            Assert.That(paragraphs, Is.Not.Empty);

            // Verify that no paragraph contains only template blocks
            foreach (var paragraph in paragraphs)
            {
                var text = paragraph.InnerText;
                // If the paragraph isn't empty, it should contain actual content, not just template syntax
                if (!string.IsNullOrWhiteSpace(text))
                {
                    Assert.That(text, Does.Not.Match(@"^\{\{.*\}\}$"),
                        "Found paragraph containing only template blocks: " + text);
                }
            }
        }
    }
}
