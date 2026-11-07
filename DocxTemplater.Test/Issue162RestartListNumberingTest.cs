using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// https://github.com/Amberg/DocxTemplater/issues/162
    /// A numbered list marked as "Restart numbering" in Word (a w:num with a w:startOverride) inside a loop
    /// must restart in every iteration of the loop - not continue across all iterations.
    /// </summary>
    internal class Issue162RestartListNumberingTest
    {
        private const int RestartNumId = 1;
        private const int ContinueNumId = 2;

        [Test]
        public void RangeLoopRestartsListInEveryIteration()
        {
            var (paragraphs, numbering) = Render(
                P("{{@i:3}}"),
                P("This Text: {{ds.Title}} - will be replaced"),
                P("{{#ds.Items}}"),
                P("{{.Name}}", RestartNumId),
                P("{{/ds.Items}}"),
                P("{{/}}"));

            AssertRestartedPerIteration(paragraphs, numbering, iterations: 3);
        }

        [Test]
        public void CollectionLoopRestartsListInEveryIteration()
        {
            var (paragraphs, numbering) = Render(
                P("{{#ds.Groups}}"),
                P("This Text: {{.Title}} - will be replaced"),
                P("{{#.Items}}"),
                P("{{.Name}}", RestartNumId),
                P("{{/.Items}}"),
                P("{{/ds.Groups}}"));

            AssertRestartedPerIteration(paragraphs, numbering, iterations: 2);
        }

        [Test]
        public void StaticListInLoopRestartsInEveryIteration()
        {
            var (paragraphs, numbering) = Render(
                P("{{@i:2}}"),
                P("Heading"),
                P("first", RestartNumId),
                P("second", RestartNumId),
                P("{{/}}"));

            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText), Is.EqualTo(["first", "second", "first", "second"]));
            Assert.That(NumId(numbered[0]), Is.EqualTo(NumId(numbered[1])));
            Assert.That(NumId(numbered[2]), Is.EqualTo(NumId(numbered[3])));
            Assert.That(NumId(numbered[0]), Is.Not.EqualTo(NumId(numbered[2])));
            AssertIsRestartOfOriginal(numbering, NumId(numbered[2]).Value);
        }

        [Test]
        public void LoopOnlyOverListItemsKeepsSingleList()
        {
            // the loop body starts with the list item itself - the items of all iterations form one list
            var (paragraphs, _) = Render(
                P("Header"),
                P("{{#ds.Items}}"),
                P("{{.Name}}", RestartNumId),
                P("{{/ds.Items}}"));

            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText), Is.EqualTo(["John", "Alice"]));
            Assert.That(numbered.Select(NumId).Distinct().Single(), Is.EqualTo(RestartNumId));
        }

        [Test]
        public void ListWithoutRestartContinuesAcrossIterations()
        {
            // no "Restart numbering" in the template - existing behavior (one continued list) is kept
            var (paragraphs, numbering) = Render(
                P("{{@i:3}}"),
                P("This Text: {{ds.Title}} - will be replaced"),
                P("{{#ds.Items}}"),
                P("{{.Name}}", ContinueNumId),
                P("{{/ds.Items}}"),
                P("{{/}}"));

            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered, Has.Count.EqualTo(6));
            Assert.That(numbered.Select(NumId).Distinct().Single(), Is.EqualTo(ContinueNumId));
            Assert.That(numbering.Elements<NumberingInstance>().Count(), Is.EqualTo(2));
        }

        private static void AssertRestartedPerIteration(List<Paragraph> paragraphs, Numbering numbering, int iterations)
        {
            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText),
                Is.EqualTo(Enumerable.Repeat(new[] { "John", "Alice" }, iterations).SelectMany(x => x)));

            var numIdsPerIteration = numbered.Chunk(2).Select(chunk => chunk.Select(NumId).Distinct().Single()).ToList();
            Assert.That(numIdsPerIteration.Distinct().Count(), Is.EqualTo(iterations), "every iteration needs its own list");
            Assert.That(numIdsPerIteration[0], Is.EqualTo(RestartNumId), "first iteration keeps the list of the template");
            foreach (var numId in numIdsPerIteration)
            {
                AssertIsRestartOfOriginal(numbering, numId.Value);
            }
        }

        private static void AssertIsRestartOfOriginal(Numbering numbering, int numId)
        {
            var original = numbering.Elements<NumberingInstance>().Single(x => x.NumberID.Value == RestartNumId);
            var instance = numbering.Elements<NumberingInstance>().Single(x => x.NumberID.Value == numId);
            Assert.That(instance.AbstractNumId.Val.Value, Is.EqualTo(original.AbstractNumId.Val.Value));
            Assert.That(instance.Elements<LevelOverride>().Single().StartOverrideNumberingValue.Val.Value, Is.EqualTo(1));
        }

        private static int? NumId(Paragraph p)
        {
            return p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
        }

        private static Paragraph P(string text, int? numId = null)
        {
            var paragraph = new Paragraph();
            if (numId != null)
            {
                paragraph.Append(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = numId.Value })));
            }
            paragraph.Append(new Run(new Text(text)));
            return paragraph;
        }

        private static (List<Paragraph> paragraphs, Numbering numbering) Render(params Paragraph[] body)
        {
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(body));
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%1." })
                        { LevelIndex = 0 })
                    { AbstractNumberId = 0 },
                    // list marked as "Restart numbering" in Word
                    new NumberingInstance(
                        new AbstractNumId { Val = 0 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
                    { NumberID = RestartNumId },
                    // plain list
                    new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = ContinueNumId });
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new
            {
                Title = "Some Text",
                Items = new[] { new { Name = "John" }, new { Name = "Alice" } },
                Groups = new[]
                {
                    new { Title = "A", Items = new[] { new { Name = "John" }, new { Name = "Alice" } } },
                    new { Title = "B", Items = new[] { new { Name = "John" }, new { Name = "Alice" } } }
                }
            });
            var result = docTemplate.Process();
            docTemplate.Validate();
            var document = WordprocessingDocument.Open(result, false);
            var mainDocumentPart = document.MainDocumentPart;
            return (mainDocumentPart.Document.Body.Elements<Paragraph>().ToList(), mainDocumentPart.NumberingDefinitionsPart.Numbering);
        }
    }
}
