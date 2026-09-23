using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class RobustnessTest
    {
        /// <summary>
        /// A schema valid table - w:tblPr and w:tblGrid are required before the first row.
        /// </summary>
        private static Table TableWith(params TableCell[] cells)
        {
            var grid = new TableGrid();
            foreach (var _ in cells)
            {
                grid.AppendChild(new GridColumn());
            }
            return new Table(new TableProperties(), grid, new TableRow(cells));
        }

        /// <summary>
        /// The block leaves the cell with content, so neither the cell nor the row around it may go.
        /// </summary>
        [Test]
        public void SingleCellTableKeepsTheContentAfterTheBlock()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                TableWith(new TableCell(
                    new Paragraph(new Run(new Text("{?{ds.A}}"))),
                    new Paragraph(new Run(new Text("{{/}}"))),
                    new Paragraph(new Run(new Text("tail")))))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { A = true });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var body = WordprocessingDocument.Open(result, false).MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(body.InnerText, Is.EqualTo("tail"));
                Assert.That(body.Descendants<TableRow>().Count(), Is.EqualTo(1));
                Assert.That(body.Descendants<TableCell>().Count(), Is.EqualTo(1));
            });
        }

        /// <summary>
        /// A cell of a row with more than one cell is never removed - that would shift the row.
        /// </summary>
        [Test]
        public void EmptiedCellOfATwoCellRowIsKept()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                TableWith(
                    new TableCell(new Paragraph(new Run(new Text("{{#ds.Items}}x{{/}}")))),
                    new TableCell(new Paragraph(new Run(new Text("keep")))))
            ));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Items = Array.Empty<string>() });
            var result = docTemplate.Process();
            docTemplate.Validate();

            var body = WordprocessingDocument.Open(result, false).MainDocumentPart.Document.Body;
            Assert.Multiple(() =>
            {
                Assert.That(body.InnerText, Is.EqualTo("keep"));
                Assert.That(body.Descendants<TableCell>().Count(), Is.EqualTo(2));
            });
        }

        [Test]
        public void ChildsBetweenReturnsTheElementsInBetween()
        {
            var first = new Paragraph(new Run(new Text("a")));
            var middle = new Paragraph(new Run(new Text("b")));
            var last = new Paragraph(new Run(new Text("c")));
            var body = new Body(first, middle, last);

            Assert.That(body.ChildsBetween(first, last), Is.EqualTo([middle]));
        }

        /// <summary>
        /// The callers remove what comes back, so the rest of the parent must not be returned when the
        /// end element does not follow the start element.
        /// </summary>
        [Test]
        public void ChildsBetweenThrowsWhenTheEndDoesNotFollowTheStart()
        {
            var first = new Paragraph(new Run(new Text("a")));
            var middle = new Paragraph(new Run(new Text("b")));
            var last = new Paragraph(new Run(new Text("c")));
            var body = new Body(first, middle, last);

            Assert.Throws<OpenXmlTemplateException>(() => body.ChildsBetween(last, first).ToList());
        }

        /// <summary>
        /// Rendering mutates the document, so a template whose rendering threw cannot be rendered or
        /// handed out again - it would be a mix of template and result.
        /// </summary>
        [Test]
        public void ProcessAfterAFailedRenderThrowsInsteadOfReturningTheTemplate()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("keep{{ds.NotThere}}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { Other = "x" });

            Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
            Assert.Multiple(() =>
            {
                Assert.That(docTemplate.Processed, Is.False);
                Assert.Throws<OpenXmlTemplateException>(() => docTemplate.Process());
                Assert.Throws<OpenXmlTemplateException>(() => docTemplate.AsStream());
            });
        }
    }
}
