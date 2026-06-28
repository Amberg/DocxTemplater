using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// Regression: when a loop marker ({{#col}} / {{/col}}) is the only content of a table cell,
    /// the marker's run was being removed and the now-empty paragraph along with it. A safety net in
    /// <c>TemplateProcessor.Cleanup</c> then re-added a bare <c>&lt;w:p/&gt;</c>, dropping the cell's
    /// intended pPr (line spacing, font size, ...). The cell ended up inheriting docDefault/Normal style,
    /// which in renderers like GemBox results in much taller rows than the template specifies.
    /// </summary>
    internal class LoopMarkerPreservesCellParagraphFormatTest
    {
        [Test]
        public void LoopMarkerOnlyCell_KeepsParagraphPropertiesAfterProcessing()
        {
            // Three-column row that gets cloned per item. Column 1 holds the loop START marker
            // and a distinctive line-spacing pPr; column 3 holds the loop END marker with its own
            // pPr. Column 2 holds the data placeholder with yet another pPr. After processing every
            // cloned row should keep all three pPrs distinguishable.
            const string content = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:tblPr><w:tblW w:w=""0"" w:type=""auto""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""1000""/><w:gridCol w:w=""1000""/><w:gridCol w:w=""1000""/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w=""1000"" w:type=""dxa""/></w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:after=""0"" w:line=""200"" w:lineRule=""exact""/></w:pPr>
        <w:r><w:t>{{#ds.items}}</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w=""1000"" w:type=""dxa""/></w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:after=""0"" w:line=""480"" w:lineRule=""auto""/></w:pPr>
        <w:r><w:t>{{.value}}</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w=""1000"" w:type=""dxa""/></w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:after=""0"" w:line=""300"" w:lineRule=""exact""/></w:pPr>
        <w:r><w:t>{{/ds.items}}</w:t></w:r>
      </w:p>
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
            docTemplate.BindModel("ds", new { items = new[] { new { value = "A" }, new { value = "B" }, new { value = "C" } } });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var body = processed.MainDocumentPart.Document.Body;

            var rows = body.Descendants<TableRow>().ToList();
            Assert.That(rows.Count, Is.EqualTo(3), "loop should produce one row per item");

            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                Assert.That(cells.Count, Is.EqualTo(3));

                // Column 1 (loop-start marker): used to become <w:p/> (no pPr) after the marker
                // paragraph was removed by RemoveWithEmptyParent and replaced by the safety net.
                // With the fix the paragraph survives with its original pPr.
                AssertCellSpacing(cells[0], expectedLine: "200", expectedRule: LineSpacingRuleValues.Exact, label: "start marker cell");

                // Column 2 (data): sanity check that the data cell is unchanged.
                AssertCellSpacing(cells[1], expectedLine: "480", expectedRule: LineSpacingRuleValues.Auto, label: "data cell");

                // Column 3 (loop-end marker): same as column 1 — pPr must survive.
                AssertCellSpacing(cells[2], expectedLine: "300", expectedRule: LineSpacingRuleValues.Exact, label: "end marker cell");
            }
        }

        [Test]
        public void RegularEmptiedParagraphInCellWithSiblingContent_IsStillRemoved()
        {
            // If a cell has another paragraph with real content alongside a marker-only paragraph,
            // the marker paragraph should still be removed (no stray empty <w:p/> left behind).
            const string content = @"
<w:tbl xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:tblPr><w:tblW w:w=""0"" w:type=""auto""/></w:tblPr>
  <w:tblGrid><w:gridCol w:w=""1000""/><w:gridCol w:w=""1000""/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w=""1000"" w:type=""dxa""/></w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:after=""0"" w:line=""200"" w:lineRule=""exact""/></w:pPr>
        <w:r><w:t>{{#ds.items}}</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>real content {{.value}}</w:t></w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w=""1000"" w:type=""dxa""/></w:tcPr>
      <w:p>
        <w:r><w:t>{{/ds.items}}</w:t></w:r>
      </w:p>
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
            docTemplate.BindModel("ds", new { items = new[] { new { value = "X" } } });
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var processed = WordprocessingDocument.Open(result, false);
            var body = processed.MainDocumentPart.Document.Body;
            var firstCell = body.Descendants<TableCell>().First();
            var paragraphs = firstCell.Elements<Paragraph>().ToList();

            // Marker paragraph removed (sibling content survives), so exactly one paragraph remains.
            Assert.That(paragraphs.Count, Is.EqualTo(1), "marker paragraph should be removed when sibling content survives in the same cell");
            Assert.That(paragraphs[0].InnerText, Is.EqualTo("real content X"));
        }

        private static void AssertCellSpacing(TableCell cell, string expectedLine, LineSpacingRuleValues expectedRule, string label)
        {
            var paragraph = cell.Elements<Paragraph>().Single();
            var ppr = paragraph.ParagraphProperties;
            Assert.That(ppr, Is.Not.Null, $"{label}: paragraph should retain its pPr");
            var spacing = ppr.Elements<SpacingBetweenLines>().SingleOrDefault();
            Assert.That(spacing, Is.Not.Null, $"{label}: spacing should survive");
            Assert.That(spacing.Line?.Value, Is.EqualTo(expectedLine), $"{label}: line value");
            Assert.That(spacing.LineRule?.Value, Is.EqualTo(expectedRule), $"{label}: line rule");
        }
    }
}
