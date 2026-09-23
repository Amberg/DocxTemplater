using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class RemoveLineBreaksAfterKeywordTest
    {
        [Test]
        public void RemoveLineBreaksAfterKeywordTestWithDocument()
        {
            using var fileStream = File.OpenRead("Resources/RemoveLineBreaksAroundSyntax.docx");
            var docTemplate = new DocxTemplate(fileStream, new ProcessSettings() { IgnoreLineBreaksAroundTags = true });
            docTemplate.BindModel("ds", new { Val = "Name", Items = new[] { "foo1", "foo2" } });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            result.Position = 0;
            var document = WordprocessingDocument.Open(result, false);
            // hash updated with the fix for https://github.com/Amberg/DocxTemplater/issues/146:
            // splitting a run before its first content no longer leaves an empty run behind
            // (one per loop iteration here), the rendered content is unchanged.
            Assert.That(TestHelper.ComputeSha256Hash(document.MainDocumentPart.Document.Body.InnerXml), Is.EqualTo("13548eda4703d0316cd6bc02e327b024533782d2e503bcd5fb47c1008883d824"));
        }


        [TestCaseSource(nameof(RemoveLineBreaksAfterKeywords_Source))]
        public void RemoveLineBreaksAfterKeywords(bool removeLineBreakEnabled, Paragraph paragraph, string expected)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(paragraph));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream, new ProcessSettings() { IgnoreLineBreaksAroundTags = removeLineBreakEnabled });
            docTemplate.BindModel("ds", "foo");
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            // check document contains newline
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            TestHelper.ExpectXmlIsEqual(body.InnerXml, expected);

        }

        private static IEnumerable<TestCaseData> RemoveLineBreaksAfterKeywords_Source()
        {
            yield return new TestCaseData(
                false,
                new Paragraph(
                    new Run(
                        new Text("Start {{.}}"),
                        new Break(),
                        new Text("{{.}} End")
                    )),
                @"
                <w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                  <w:r>
                    <w:t xml:space=""preserve"">Start </w:t>
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:br />
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:t xml:space=""preserve""> End</w:t>
                  </w:r>
                </w:p>")
                .SetName("Not Enabled");

            yield return new TestCaseData(
                true,
                new Paragraph(
                    new Run(
                        new Text("Start {{.}}"),
                        new Break(),
                        new Text("{{.}} End")
                    )),
                @"
                <w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                  <w:r>
                    <w:t xml:space=""preserve"">Start </w:t>
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:t xml:space=""preserve""> End</w:t>
                  </w:r>
                </w:p>")
                .SetName("Enabled br removed between");

            yield return new TestCaseData(
                true,
                new Paragraph(
                    new Run(
                        new Text("Start {{.}} Text"),
                        new Break(),
                        new Text("Text {{.}} End")
                    )),
                @"
                <w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                  <w:r>
                    <w:t xml:space=""preserve"">Start </w:t>
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:t xml:space=""preserve""> Text</w:t>
                    <w:br />
                    <w:t xml:space=""preserve"">Text </w:t>
                    <w:t xml:space=""preserve"">foo</w:t>
                    <w:t xml:space=""preserve""> End</w:t>
                  </w:r>
                </w:p>")
                .SetName("Enabled but not removed because of text");
        }
    }
}
