using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// Regression: a multi-line value (or a value with a trailing newline) bound into a text node
    /// must NOT produce a trailing &lt;w:br/&gt;. A trailing break renders as an empty line at the
    /// bottom of the containing paragraph/cell - in a table that looks like an extra blank row,
    /// which a renderer like GemBox then shows as a visible gap before the next row.
    /// </summary>
    internal class MultiLineValueInCellNoTrailingBreakTest
    {
        [Test]
        public void MultiLineValue_HasBreaksBetweenLinesButNoTrailingBreak()
        {
            const string content = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:tblPr><w:tblW w:w=""0"" w:type=""auto""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""5000""/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w=""5000"" w:type=""dxa""/></w:tcPr>
      <w:p><w:r><w:t>{{ds.Description}}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>";

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document { Body = new Body { InnerXml = content } };
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Description = "Line one\nLine two\nLine three" });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var cell = processed.MainDocumentPart.Document.Body.Descendants<TableCell>().Single();
            var paragraph = cell.Elements<Paragraph>().Single();

            // Three text lines -> exactly two breaks (between them), never a trailing one.
            var texts = paragraph.Descendants<Text>().Select(t => t.Text).ToList();
            var breaks = paragraph.Descendants<Break>().Count();
            Assert.That(texts, Is.EqualTo(["Line one", "Line two", "Line three"]));
            Assert.That(breaks, Is.EqualTo(2), "n lines must yield n-1 breaks, with no trailing break");

            // The last child element of the run must be the text, not a break.
            var run = paragraph.Elements<Run>().Last();
            Assert.That(run.LastChild, Is.TypeOf<Text>(), "run must not end with a trailing break");
        }

        [Test]
        public void BreakCountEqualsNewlineCount()
        {
            // The rule is exactly one break per '\n'. A blank line in the middle must be preserved,
            // and a genuine trailing newline still yields its break (see EachItemOnNewLine) - what
            // must NOT happen is an extra break with no corresponding '\n'.
            using (var processed = Render("Line one\n\nLine three"))
            {
                var paragraph = processed.MainDocumentPart.Document.Body.Descendants<TableCell>().Single().Elements<Paragraph>().Single();
                Assert.That(paragraph.Descendants<Text>().Select(t => t.Text), Is.EqualTo(["Line one", "Line three"]));
                Assert.That(paragraph.Descendants<Break>().Count(), Is.EqualTo(2), "blank middle line must be preserved as two breaks");
            }

            using (var processed = Render("Trailing\n"))
            {
                var paragraph = processed.MainDocumentPart.Document.Body.Descendants<TableCell>().Single().Elements<Paragraph>().Single();
                Assert.That(paragraph.Descendants<Break>().Count(), Is.EqualTo(1), "a real trailing newline maps to exactly one break");
            }
        }

        private static WordprocessingDocument Render(string description)
        {
            const string content = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:tblPr><w:tblW w:w=""0"" w:type=""auto""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""5000""/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w=""5000"" w:type=""dxa""/></w:tcPr>
      <w:p><w:r><w:t>{{ds.Description}}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>";

            var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document { Body = new Body { InnerXml = content } };
                wpDocument.Save();
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Description = description });
            var result = docTemplate.Process();
            docTemplate.Validate();
            return WordprocessingDocument.Open(result, false);
        }
    }
}
