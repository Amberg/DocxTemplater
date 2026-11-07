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
        private const int SubListRestartNumId = 3;

        [TestCase("{{@i:3}}", "{{ds.Title}}", "ds.Items", "{{/}}", 3)]
        [TestCase("{{#ds.Groups}}", "{{.Title}}", ".Items", "{{/ds.Groups}}", 2)]
        public void LoopRestartsListInEveryIteration(string loopStart, string title, string items, string loopEnd, int iterations)
        {
            var (paragraphs, numbering) = Render(
                P(loopStart),
                P($"This Text: {title} - will be replaced"),
                P($"{{{{#{items}}}}}"),
                P("{{.Name}}", RestartNumId),
                P($"{{{{/{items}}}}}"),
                P(loopEnd));

            AssertRestartedPerIteration(paragraphs, numbering, ["John", "Alice"], iterations);
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

            AssertRestartedPerIteration(paragraphs, numbering, ["first", "second"], 2);
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
        public void InlineLoopInListParagraphKeepsSingleList()
        {
            var (paragraphs, _) = Render(
                P("{{#ds.Items}}{{.Name}} {{/ds.Items}}", RestartNumId));

            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText.Trim()), Is.EqualTo(["John Alice"]));
            Assert.That(NumId(numbered[0]), Is.EqualTo(RestartNumId));
        }

        [Test]
        public void LoopWritesListWithRestartedSubListPerItem()
        {
            // the loop writes the items of the outer list - every item has a sub list restarting at 1
            var (paragraphs, _) = Render(
                P("{{#ds.Groups}}"),
                P("{{.Title}}", ContinueNumId),
                P("{{#.Items}}"),
                P("{{.Name}}", SubListRestartNumId),
                P("{{/.Items}}"),
                P("{{/ds.Groups}}"));

            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText), Is.EqualTo(["A", "John", "Alice", "B", "John", "Alice"]));
            // outer list continues over the groups
            Assert.That(NumId(numbered[0]), Is.EqualTo(ContinueNumId));
            Assert.That(NumId(numbered[3]), Is.EqualTo(ContinueNumId));
            // sub list: one list per group
            Assert.That(NumId(numbered[1]), Is.EqualTo(NumId(numbered[2])));
            Assert.That(NumId(numbered[4]), Is.EqualTo(NumId(numbered[5])));
            Assert.That(NumId(numbered[1]), Is.Not.EqualTo(NumId(numbered[4])));
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
            Assert.That(numbering.Elements<NumberingInstance>().Count(), Is.EqualTo(3));
        }

        private static void AssertRestartedPerIteration(List<Paragraph> paragraphs, Numbering numbering, string[] items, int iterations)
        {
            var numbered = paragraphs.Where(x => NumId(x) != null).ToList();
            Assert.That(numbered.Select(x => x.InnerText),
                Is.EqualTo(Enumerable.Repeat(items, iterations).SelectMany(x => x)));

            var numIdsPerIteration = numbered.Chunk(items.Length).Select(chunk => chunk.Select(NumId).Distinct().Single()).ToList();
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
                    new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = ContinueNumId },
                    // second list marked as "Restart numbering" - used as sub list
                    new NumberingInstance(
                        new AbstractNumId { Val = 0 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 })
                    { NumberID = SubListRestartNumId });
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
