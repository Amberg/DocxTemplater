using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    internal class BookmarkInDifferentTableRowDoesNotcauseCrash
    {
        [Test]
        public void Test()
        {
            string content = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">  
                              <w:r>  
                                <w:t xml:space=""preserve"">This is sentence one.</w:t>  
                              </w:r>  
                              <w:bookmarkStart w:id=""0"" w:name=""testing123""/>  
                              <w:r>  
                                <w:t>This is sentence two. {{.}}</w:t>  
                              </w:r>  
                            </w:p>  
                            <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">  
                              <w:r>  
                                <w:t xml:space=""preserve"">This </w:t>  
                              </w:r>  
                              <w:bookmarkEnd w:id=""0""/>  
                              <w:r>  
                                <w:t>is sentence three.{{.}}</w:t>  
                              </w:r>  
                            </w:p>";

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            var mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document
            {
                Body = new Body
                {
                    InnerXml = content
                }
            };
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", "hi there");
            var result = docTemplate.Process();
            docTemplate.Validate();

            // get body
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.Descendants<BookmarkEnd>().Count(), Is.EqualTo(0));
            Assert.That(body.Descendants<BookmarkStart>().Count(), Is.EqualTo(0));
        }
    }
}
