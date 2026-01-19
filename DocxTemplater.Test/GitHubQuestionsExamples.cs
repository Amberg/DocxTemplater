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
                    @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:r><w:t>First Line</w:t><w:br /><w:t>Second Line</w:t><w:br /><w:t>Third Line</w:t><w:br /></w:r></w:p>"));
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
