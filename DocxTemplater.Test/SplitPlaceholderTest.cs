using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    class SplitPlaceholderTest
    {
        [Test]
        public void ProcessDocWithSplitPlaceholders()
        {
            using var fileStream = File.OpenRead("Resources/SplitPlaceholdersTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.BindModel("ds", new { MyPlaceholder = " !!!!!!!! ", Last = "LAST" });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();

            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var runTexts = document.MainDocumentPart.Document.Body.Descendants<Run>()
                .SelectMany(x => x.Descendants<Text>().Select(x => x.Text));
            Assert.That(runTexts,
                Is.EqualTo(new[]
                    {"AAAAA", " !!!!!!!! ", "BB", "BBB", " !!!!!!!! ", "CC", "CCC", " !!!!!!!! ", "LAST", "DDDD"}));
        }

        [Test]
        public void SimpleVariableSplit_FirstVariableLongerAsSplittedPart()
        {
            var template = "Test:{{.variable}}{{.variable}}";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(template)))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("Items", new { variable = "foo" });
            var result = docTemplate.Process();
            result.Position = 0;
            // compare body
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerXml, Is.EqualTo("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                                  "<w:r>" +
                                                  "<w:t xml:space=\"preserve\">Test:</w:t>" +
                                                  "<w:t xml:space=\"preserve\">foo</w:t>" +
                                                  "<w:t xml:space=\"preserve\">foo</w:t>" +
                                                  "</w:r>" +
                                                  "</w:p>"));
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

            var characterMap = new CharacterMap(paragraph);
            var firstChar = characterMap[characterMap.Text.IndexOf("Text1")];
            var lastChar = characterMap[characterMap.Text.IndexOf("Text4") + 4];
            characterMap.MergeText(firstChar, lastChar);

            Console.WriteLine(paragraph.ToPrettyPrintXml());
            texts = paragraph.Descendants<Text>().ToList();
            Assert.That(texts.Count, Is.EqualTo(5));
            Assert.That(texts[0].Text, Is.EqualTo("Leading Text Same Run"));
            Assert.That(texts[1].Text, Is.EqualTo("##"));
            Assert.That(texts[2].Text, Is.EqualTo("Text1Text2Text3Text4"));
            Assert.That(texts[3].Text, Is.EqualTo("##NotPart"));
            Assert.That(texts[4].Text, Is.EqualTo("Not Merged after"));
        }

        [TestCase("None", new[] { "None" })]
        [TestCase("leading{{Var}}", new[] { "leading", "{{Var}}" })]
        [TestCase("{{Var}}trailing", new[] { "{{Var}}", "trailing" })]
        [TestCase("12345678{{Var}}trailing", new[] { "12345678", "{{Var}}", "trailing" })]
        [TestCase("1{{Var}}t{{Var}}a", new[] { "1", "{{Var}}", "t", "{{Var}}", "a" })]

        public void MergeText(string text, string[] expected)
        {
            var xml = $@"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <w:r>
                        <w:t>{text}</w:t> 
                        </w:r>             
                    </w:p>";
            var paragraph = new Paragraph(xml);
            var characterMap = new CharacterMap(paragraph);
            foreach (var m in PatternMatcher.FindSyntaxPatterns(characterMap.Text))
            {
                var firstChar = characterMap[m.Index];
                var lastChar = characterMap[m.Index + m.Length - 1];
                characterMap.MergeText(firstChar, lastChar);
            }

            Assert.That(paragraph.Descendants<Text>().Select(x => x.Text), Is.EqualTo(expected));
        }
    }
}

