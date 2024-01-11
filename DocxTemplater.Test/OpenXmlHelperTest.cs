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
            Assert.True(splitedParts.All(x => x is Paragraph));
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
            Assert.True(splitedParts.All(x => x is Paragraph));
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
            Assert.True(splitedParts.All(x => x is Run));
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
            Assert.True(splitedParts.All(x => x is Paragraph));
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
            Assert.True(splitedParts.All(x => x is Run));
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
            Assert.True(splitedParts.All(x => x is Run));
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
            Assert.True(splitedParts.All(x => x is Run));
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
            CollectionAssert.AreEqual(parts.Select(x => x.InnerText), new[] { "Text1", "Text2" });

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
            CollectionAssert.AreEqual(parts.Select(x => x.InnerText), new[] { "Text1" });

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

        [Test]
        public void MergeTextToOneRun()
        {
            var xml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <w:r>
                        <w:t>Leading Text Same Run</w:t>
                        <w:t>##Text1</w:t>
                        </w:r>
                        <w:r>
                        <w:t>Text2</w:t>
                        <w:t>Text3</w:t>
                        </w:r>
                        <w:r>
                            <w:t>Text4##NotPart</w:t>
                        </w:r>
                        <w:r>
                            <w:t>Not Merged after</w:t>
                        </w:r>
                    </w:p>";
            var paragraph = new Paragraph(xml);
            var texts = paragraph.Descendants<Text>().ToList();
            var firstText = texts.Single(x => x.Text.EndsWith("Text1"));

            firstText.MergeText(2, texts.Single(x => x.Text.StartsWith("Text4")), 20);

            Console.WriteLine(paragraph.ToPrettyPrintXml());
            texts = paragraph.Descendants<Text>().ToList();
            Assert.That(texts.Count, Is.EqualTo(5));
            Assert.That(texts[0].Text, Is.EqualTo("Leading Text Same Run"));
            Assert.That(texts[1].Text, Is.EqualTo("##"));
            Assert.That(texts[2].Text, Is.EqualTo("Text1Text2Text3Text4"));
            Assert.That(texts[3].Text, Is.EqualTo("##NotPart"));
            Assert.That(texts[4].Text, Is.EqualTo("Not Merged after"));
        }

        [Test]
        public void MergeTextMiddle()
        {
            var xml = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <w:r>
                        <w:t>LeadingMiddleEnd</w:t> 
                        </w:r>             
                    </w:p>";
            var paragraph = new Paragraph(xml);
            var firstText = paragraph.Descendants<Text>().Single(x => x.Text == "LeadingMiddleEnd");
            firstText.MergeText(7, firstText, 6);
            CollectionAssert.AreEqual(new[] { "Leading", "Middle", "End" }, paragraph.Descendants<Text>().Select(x => x.Text));
        }
    }
}
