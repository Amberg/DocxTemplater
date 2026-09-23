using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Markdown;

namespace DocxTemplater.Test
{
    class GitHubQuestionsExamples
    {

        [Test]
        public void EachItemOnNewLine()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{#Items}}{{.}}{{/Items}}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("Items", new[] { "First Line\r\n", "Second Line\r\n", "Third Line\r\n" });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
            // check document contains 2 altChunks
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.InnerXml,
                Is.EqualTo(
                    @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:t xml:space=""preserve"">First Line</w:t><w:br /><w:t xml:space=""preserve"">Second Line</w:t><w:br /><w:t xml:space=""preserve"">Third Line</w:t><w:br /></w:r></w:p>"));
        }



        [Test]
        // Issue: https://github.com/Amberg/DocxTemplater/issues/114
        public void InsertionPointConditionNotFound_Issue114()
        {
            using var fileStream = File.OpenRead("Resources/Issue114.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var data = new Dictionary<string, object>
            {
                { "RESIDES", 0 },
                { "CATEGORY", "cat_002" },
                { "DEMAND_TYPE_CODE", "00017" }
            };
            docTemplate.BindModel("ds", data);
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();
        }

        // Regression for https://github.com/Amberg/DocxTemplater/issues/114
        // Trigger pattern: nested condition (with else) spanning paragraphs whose start text
        // sits in a paragraph carrying the inherited end-marker of the surrounding block,
        // followed by an independent top-level condition. The old SplitAtElement cloned the
        // IpId attribute, multiple paragraphs ended up sharing the same insertion-point id,
        // and the next-sibling skip in AddInsertionPoints placed later anchors far past
        // their intended location, corrupting the trailing top-level block.
        [Test]
        public void NestedConditionsAcrossParagraphsWithSiblingBlock_Issue114()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{?{ds.O1}}{?{ds.O2}}"))),
                new Paragraph(new Run(new Text("A{{:}}{?{ds.I}}B{{:}}C"))),
                new Paragraph(new Run(new Text("{{/}}{{/}}{{/}}"))),
                new Paragraph(new Run(new Text("ok"))),
                new Paragraph(new Run(new Text("{?{ds.X}}"))),
                new Paragraph(new Run(new Text("Y"))),
                new Paragraph(new Run(new Text("{{/}}")))
            ));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new { O1 = true, O2 = false, I = true, X = true });
            var result = docTemplate.Process();
            docTemplate.Validate();
            var document = WordprocessingDocument.Open(result, false);
            var text = document.MainDocumentPart.Document.Body.InnerText;
            Assert.That(text, Is.EqualTo("BokY"));
        }

        // Issue https://github.com/Amberg/DocxTemplater/issues/146
        // A block boundary that falls between two runs splits the run holding the start tag.
        // Both halves have to keep the run properties, otherwise the content after the boundary
        // falls back to docDefaults instead of the formatting of the template.
        [Test]
        public void BlockBoundaryInsideRunKeepsRunFormatting_Issue146()
        {
            var body = ProcessBodyTemplate(
                @"<w:p>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{?{ds.Cond}}{{ds.A}}</w:t></w:r>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{{/}}{{ds.B}}</w:t></w:r>
                  </w:p>",
                new { Cond = true, A = "foo", B = "bar" });

            Assert.That(body.InnerText, Is.EqualTo("foobar"));
            var runs = body.Descendants<Run>().Where(r => r.InnerText.Length > 0).ToList();
            Assert.That(runs.Select(r => r.InnerText), Is.EqualTo(["foo", "bar"]));
            Assert.That(runs.Select(r => r.RunProperties?.Bold), Has.All.Not.Null,
                "both halves of the split run must keep the bold run properties");
        }

        // The block spans two paragraphs, so the paragraphs are split as well and both halves
        // have to keep the paragraph properties.
        [Test]
        public void BlockBoundaryAcrossParagraphsKeepsParagraphFormatting_Issue146()
        {
            var body = ProcessBodyTemplate(
                @"<w:p>
                    <w:pPr><w:jc w:val=""center"" /></w:pPr>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{?{ds.Cond}}</w:t></w:r>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{{ds.A}}</w:t></w:r>
                  </w:p>
                  <w:p>
                    <w:pPr><w:jc w:val=""center"" /></w:pPr>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{{/}}</w:t></w:r>
                    <w:r><w:rPr><w:b /></w:rPr><w:t>{{ds.B}}</w:t></w:r>
                  </w:p>",
                new { Cond = true, A = "foo", B = "bar" });

            Assert.That(body.InnerText, Is.EqualTo("foobar"));
            var paragraphs = body.Descendants<Paragraph>().Where(p => p.InnerText.Length > 0).ToList();
            Assert.That(paragraphs.Select(p => p.InnerText), Is.EqualTo(["foo", "bar"]));
            Assert.That(paragraphs.Select(p => p.ParagraphProperties?.Justification?.Val?.Value),
                Has.All.EqualTo(JustificationValues.Center),
                "both halves of the split paragraph must keep the centered justification");
            Assert.That(body.Descendants<Run>().Where(r => r.InnerText.Length > 0).Select(r => r.RunProperties?.Bold),
                Has.All.Not.Null);
        }

        // Properties that exist once per paragraph must not be copied to both halves - a duplicated
        // w:sectPr adds a section break, a duplicated w:numPr an extra list number, and a duplicated
        // tracked change record breaks the document unique w:id.
        [Test]
        public void SplitDoesNotDuplicateSingleParagraphProperties_Issue146()
        {
            var body = ProcessBodyTemplate(
                @"<w:p>
                    <w:pPr>
                      <w:pageBreakBefore />
                      <w:numPr><w:ilvl w:val=""0"" /><w:numId w:val=""1"" /></w:numPr>
                      <w:rPr><w:ins w:id=""7"" w:author=""a"" w:date=""2020-01-01T00:00:00Z"" /></w:rPr>
                    </w:pPr>
                    <w:r><w:t>Head {?{ds.Cond}}</w:t></w:r>
                    <w:r><w:t>Tail</w:t></w:r>
                  </w:p>
                  <w:p><w:r><w:t>{{/}}</w:t></w:r></w:p>",
                new { Cond = true });

            Assert.That(body.Descendants<NumberingProperties>().Count(), Is.EqualTo(1));
            Assert.That(body.Descendants<PageBreakBefore>().Count(), Is.EqualTo(1));
            Assert.That(body.Descendants<Inserted>().Select(x => x.Id?.ToString()), Is.EqualTo(["7"]),
                "the tracked change id has to stay unique in the document");
        }

        [Test]
        public void SplitDoesNotDuplicateSectionProperties_Issue146()
        {
            var body = ProcessBodyTemplate(
                @"<w:p>
                    <w:pPr><w:sectPr><w:pgSz w:w=""11906"" w:h=""16838"" /></w:sectPr></w:pPr>
                    <w:r><w:t>Head {?{ds.Cond}}</w:t></w:r>
                    <w:r><w:t>Tail</w:t></w:r>
                  </w:p>
                  <w:p><w:r><w:t>{{/}}</w:t></w:r></w:p>",
                new { Cond = true });

            Assert.That(body.Descendants<SectionProperties>().Count(), Is.EqualTo(1));
        }

        private static Body ProcessBodyTemplate(string bodyContentXml, object model)
        {
            const string wordprocessingNamespace = @"xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";
            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body($"<w:body {wordprocessingNamespace}>{bodyContentXml}</w:body>"));
                wpDocument.Save();
            }
            memStream.Position = 0;

            using var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            using var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body;
        }

        [Test]
        // Issue: https://github.com/Amberg/DocxTemplater/issues/89
        public void MarkdownCrash()
        {
            using var fileStream = File.OpenRead("Resources/markdown-crash-example.docx");
            var docTemplate = new DocxTemplate(fileStream);
            var data = new
            {
                var_hastegenargumenten = true,
                var_tegenargumenten = new[]
                {
                    new
                    {
                        nrq_name = "Tegenargument",
                        nrq_argument = "BASIS ARGUMENT",
                        nrq_response = "SOME RESPONSE",
                        nrq_hidedecisiononletter = new
                        {
                            Label = "Nee",
                            Value = false
                        },
                        nrq_decision = new
                        {
                            Label = "Niet geaccepteerd",
                            Value = 875810001
                        },
                        nrq_regarding = new
                        {
                            Label = "Dossier",
                            Value = 875810000
                        },
                        nrq_processphase = new
                        {
                            Label = "Hoorzitting",
                            Value = 875810003
                        }
                    }
                }
            };
            docTemplate.RegisterFormatter(new MarkdownFormatter());
            docTemplate.BindModel("ds", data);
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.SaveAsFileAndOpenInWord();

        }
    }
}
