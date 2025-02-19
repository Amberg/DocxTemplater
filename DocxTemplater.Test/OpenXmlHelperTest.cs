using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class OpenXmlHelperTest
    {
        [Test]
        public void SplitAfterElementSameRunAtParagraphLevel()
        {
            var paragraph = new Paragraph(new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail")));
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "StartSplit");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = paragraph.SplitAfterElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(2));
            Assert.That(splitedParts.All(x => x is Paragraph));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(2));
            Assert.That(splitedParts.ElementAt(1).Descendants<Text>().Count(), Is.EqualTo(3));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }

        [Test]
        public void SplitAfterElementDifferentRunAtParagraphLevel()
        {
            var paragraph = new Paragraph(
                new Run(new Text("Leading")),
                new Run(new Text("StartSplit")),
                new Run(new Text("SplitContent")),
                new Run(new Text("FirstAfterSplit"), new Text("Trail")));
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "StartSplit");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = paragraph.SplitAfterElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(2));
            Assert.That(splitedParts.All(x => x is Paragraph));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(2));
            Assert.That(splitedParts.ElementAt(1).Descendants<Text>().Count(), Is.EqualTo(3));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }


        [Test]
        public void SplitAfterElemntSameRunAtRunLevel()
        {
            var run = new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail"));
            var paragraph = new Paragraph(run);
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "StartSplit");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = run.SplitAfterElement(element);
            Assert.That(splitedParts.Count, Is.EqualTo(2));
            Assert.That(splitedParts.All(x => x is Run));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(2));
            Assert.That(splitedParts.ElementAt(1).Descendants<Text>().Count(), Is.EqualTo(3));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }

        [Test]
        public void SplitBeforeElementSameRunAtParagraphLevel()
        {
            var paragraph = new Paragraph(new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail")));
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "StartSplit");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = paragraph.SplitBeforeElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(2));
            Assert.That(splitedParts.All(x => x is Paragraph));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(1));
            Assert.That(splitedParts.ElementAt(1).Descendants<Text>().Count(), Is.EqualTo(4));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }

        [Test]
        public void SplitBeforeElementSameRunAtRunLevel()
        {
            var run = new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail"));
            var paragraph = new Paragraph(run);
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "StartSplit");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = run.SplitBeforeElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(2));
            Assert.That(splitedParts.All(x => x is Run));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(1));
            Assert.That(splitedParts.ElementAt(1).Descendants<Text>().Count(), Is.EqualTo(4));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }


        [Test]
        public void SplitAfterElementAtEndOfParent()
        {
            var run = new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail"));
            var paragraph = new Paragraph(run);
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "Trail");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = run.SplitAfterElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(1));
            Assert.That(splitedParts.All(x => x is Run));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(5));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }


        [Test]
        public void SplitBeforeElementAtStartOfParent()
        {
            var run = new Run(new Text("Leading"), new Text("StartSplit"), new Text("SplitContent"),
                new Text("FirstAfterSplit"), new Text("Trail"));
            var paragraph = new Paragraph(run);
            var body = new Document(new Body(paragraph));
            var element = body.Descendants<Text>().Single(x => x.Text == "Leading");
            Console.WriteLine(body.ToPrettyPrintXml());

            var innerTextBefore = body.InnerText;

            var splitedParts = run.SplitBeforeElement(element);

            Assert.That(splitedParts.Count, Is.EqualTo(1));
            Assert.That(splitedParts.All(x => x is Run));
            Assert.That(splitedParts.ElementAt(0).Descendants<Text>().Count(), Is.EqualTo(5));
            Assert.That(body.InnerText, Is.EqualTo(innerTextBefore));
        }

        [Test]
        public void SplitRun()
        {
            var xml = @"
                         <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                           <w:r>
                            <w:t>Text1</w:t>
                            <w:t>Text2</w:t>
                           </w:r>
                        </w:p>
                        ";
            var paragraph = new Paragraph(xml);
            var parts = paragraph.ChildElements.First<Run>().SplitAfterElement(paragraph.Descendants<Text>().First());
            Assert.That(parts.Count, Is.EqualTo(2));
            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(new[] { "Text1", "Text2" }));

            var runs = paragraph.Descendants<Run>().ToList();
            Assert.That(runs.Count, Is.EqualTo(2));
            Assert.That(runs[0].InnerText, Is.EqualTo("Text1"));
            Assert.That(runs[1].InnerText, Is.EqualTo("Text2"));
        }

        [Test]
        public void SplitRunOnlyOneTextInRun()
        {
            var xml = @"
                         <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                           <w:r>
                            <w:t>Text1</w:t>
                           </w:r>
                        </w:p>
                        ";
            var paragraph = new Paragraph(xml);
            var parts = paragraph.ChildElements.First<Run>().SplitAfterElement(paragraph.Descendants<Text>().First());
            Assert.That(parts.Count, Is.EqualTo(1));
            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(new[] { "Text1" }));

            var runs = paragraph.Descendants<Run>().ToList();
            Assert.That(runs.Count, Is.EqualTo(1));
            Assert.That(runs[0].InnerText, Is.EqualTo("Text1"));
        }


        [Test]
        public void SplitBeforeSplitMarkerIsLastElement()
        {
            var xml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <w:r>
                        <w:t>{{ds.Items.Value}}</w:t>
                        </w:r>
                        <w:r>
                        <w:t/>
                        </w:r>
                    </w:p>";
            var paragraph = new Paragraph(xml);
            var lastText = paragraph.Descendants<Text>().Last();
            var parts = paragraph.SplitBeforeElement(lastText);
            Assert.That(parts.Count, Is.EqualTo(2));
            Assert.That(parts.ElementAt(0).InnerText, Is.EqualTo("{{ds.Items.Value}}"));
            Assert.That(parts.ElementAt(1).InnerText, Is.EqualTo(string.Empty));
        }
    }
}
