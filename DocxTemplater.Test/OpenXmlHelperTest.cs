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
        public void SplitAfterElementSameRunAtRunLevel()
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
            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1", "Text2"]));

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
            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1"]));

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

        // Issue https://github.com/Amberg/DocxTemplater/issues/146
        // SplitAtElement shallow-clones the parent, so the formatting properties (w:rPr / w:pPr)
        // were missing on the clone and its content fell back to docDefaults.
        [Test]
        public void SplitAfterElementKeepsRunPropertiesOnBothParts()
        {
            var body = ParseBody(@"<w:p>
                                     <w:r>
                                       <w:rPr><w:b /></w:rPr>
                                       <w:t>Text1</w:t>
                                       <w:t>Text2</w:t>
                                     </w:r>
                                   </w:p>");
            var parts = body.Descendants<Run>().First().SplitAfterElement(body.Descendants<Text>().First());

            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1", "Text2"]));
            Assert.That(parts.Cast<Run>().Select(x => x.RunProperties?.Bold), Has.All.Not.Null);
            Assert.That(body.InnerText, Is.EqualTo("Text1Text2"));
        }

        [Test]
        public void SplitBeforeElementKeepsRunPropertiesOnBothParts()
        {
            var body = ParseBody(@"<w:p>
                                     <w:r>
                                       <w:rPr><w:b /></w:rPr>
                                       <w:t>Text1</w:t>
                                       <w:t>Text2</w:t>
                                     </w:r>
                                   </w:p>");
            var parts = body.Descendants<Run>().First().SplitBeforeElement(body.Descendants<Text>().Last());

            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1", "Text2"]));
            Assert.That(parts.Cast<Run>().Select(x => x.RunProperties?.Bold), Has.All.Not.Null);
            Assert.That(body.InnerText, Is.EqualTo("Text1Text2"));
        }

        [Test]
        public void SplitKeepsParagraphPropertiesOnBothParts()
        {
            var body = ParseBody(@"<w:p>
                                     <w:pPr><w:jc w:val=""center"" /></w:pPr>
                                     <w:r><w:rPr><w:b /></w:rPr><w:t>Text1</w:t></w:r>
                                     <w:r><w:rPr><w:b /></w:rPr><w:t>Text2</w:t></w:r>
                                   </w:p>");
            var parts = body.Descendants<Paragraph>().First().SplitAfterElement(body.Descendants<Text>().First());

            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1", "Text2"]));
            Assert.That(parts.Cast<Paragraph>().Select(x => x.ParagraphProperties?.Justification?.Val?.Value),
                Has.All.EqualTo(JustificationValues.Center));
            Assert.That(parts.SelectMany(x => x.Descendants<Run>()).Select(x => x.RunProperties?.Bold), Has.All.Not.Null);
            Assert.That(body.InnerText, Is.EqualTo("Text1Text2"));
        }

        // The paragraph mark properties and the tracked change records are not duplicated - their
        // w:id has to stay unique in the document - they stay on the first of the two halves.
        [Test]
        public void SplitDoesNotDuplicateParagraphMarkProperties()
        {
            var body = ParseBody(@"<w:p>
                                     <w:pPr>
                                       <w:jc w:val=""center"" />
                                       <w:numPr><w:ilvl w:val=""0"" /><w:numId w:val=""1"" /></w:numPr>
                                       <w:rPr><w:ins w:id=""7"" w:author=""a"" w:date=""2020-01-01T00:00:00Z"" /></w:rPr>
                                     </w:pPr>
                                     <w:r><w:t>Text1</w:t></w:r>
                                     <w:r><w:t>Text2</w:t></w:r>
                                   </w:p>");
            var parts = body.Descendants<Paragraph>().First().SplitAfterElement(body.Descendants<Text>().First());

            Assert.That(parts.Cast<Paragraph>().Select(x => x.ParagraphProperties?.Justification?.Val?.Value),
                Has.All.EqualTo(JustificationValues.Center), "formatting is copied to both halves");
            Assert.That(body.Descendants<NumberingProperties>().Count(), Is.EqualTo(1));
            Assert.That(body.Descendants<ParagraphMarkRunProperties>().Count(), Is.EqualTo(1));
            Assert.That(body.Descendants<Inserted>().Count(), Is.EqualTo(1));
            // they belong to the first half, as before the split
            Assert.That(parts.First().Descendants<NumberingProperties>().Count(), Is.EqualTo(1));
        }

        [Test]
        public void SplitBeforeElementKeepsParagraphMarkPropertiesOnFirstPart()
        {
            var body = ParseBody(@"<w:p>
                                     <w:pPr>
                                       <w:numPr><w:ilvl w:val=""0"" /><w:numId w:val=""1"" /></w:numPr>
                                     </w:pPr>
                                     <w:r><w:t>Text1</w:t></w:r>
                                     <w:r><w:t>Text2</w:t></w:r>
                                   </w:p>");
            var parts = body.Descendants<Paragraph>().First().SplitBeforeElement(body.Descendants<Text>().Last());

            Assert.That(body.Descendants<NumberingProperties>().Count(), Is.EqualTo(1));
            Assert.That(parts.First().Descendants<NumberingProperties>().Count(), Is.EqualTo(1));
        }

        // A table keeps its w:tblPr and the REQUIRED w:tblGrid on both halves - w:bookmarkStart may
        // legally precede them, the properties are not simply the first children.
        [TestCase("")]
        [TestCase(@"<w:bookmarkStart w:id=""1"" w:name=""bm"" />")]
        public void SplitTableKeepsTablePropertiesOnBothParts(string leadingRangeMarkup)
        {
            var body = ParseBody($@"<w:tbl>
                                     {leadingRangeMarkup}
                                     <w:tblPr><w:tblStyle w:val=""TableGrid"" /></w:tblPr>
                                     <w:tblGrid><w:gridCol w:w=""100"" /></w:tblGrid>
                                     <w:tr><w:trPr><w:cantSplit /></w:trPr><w:tc><w:tcPr /><w:p><w:r><w:t>Text1</w:t></w:r></w:p></w:tc></w:tr>
                                     <w:tr><w:trPr><w:cantSplit /></w:trPr><w:tc><w:tcPr /><w:p><w:r><w:t>Text2</w:t></w:r></w:p></w:tc></w:tr>
                                   </w:tbl>");
            var table = body.Descendants<Table>().First();
            var parts = table.SplitAfterElement(body.Descendants<TableRow>().First());

            Assert.That(parts.Select(x => x.InnerText), Is.EqualTo(["Text1", "Text2"]));
            Assert.That(parts.Cast<Table>().Select(x => x.GetFirstChild<TableProperties>()), Has.All.Not.Null);
            Assert.That(parts.Cast<Table>().Select(x => x.GetFirstChild<TableGrid>()), Has.All.Not.Null);
            Assert.That(body.InnerText, Is.EqualTo("Text1Text2"));
            body.ValidateOpenXmlElement();
        }

        private static Body ParseBody(string bodyContentXml)
        {
            const string wordprocessingNamespace = @"xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";
            return new Body($"<w:body {wordprocessingNamespace}>{bodyContentXml}</w:body>");
        }
    }
}
