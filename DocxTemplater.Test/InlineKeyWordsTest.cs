using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplater.Test
{
    class InlineKeyWordsTest
    {
        [TestCase("pageBreak", "<w:br w:type=\"page\" />")]
        [TestCase("break", "<w:br />")]
        [TestCase("BrEaK", "<w:br />")]
        public void InsertSimpleBreaks(string keyWord, string xmlElement)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text($$$"""Here comes a break{{:{{{keyWord}}}}}and here comes the second run""")))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.ToPrettyPrintXml(), Is.EqualTo($"""
                                                              <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                                                <w:p>
                                                                  <w:r>
                                                                    <w:t xml:space="preserve">Here comes a break</w:t>
                                                                    {xmlElement}
                                                                    <w:t xml:space="preserve">and here comes the second run</w:t>
                                                                  </w:r>
                                                                </w:p>
                                                              </w:body>
                                                              """));
        }

        [Test]
        public void InsertSectionBreak()
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text($$$"""Here comes a break{{:sectionBreak}}and here comes the second part""")))));
            wpDocument.Save();
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);

            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;
            Assert.That(body.ToPrettyPrintXml(), Is.EqualTo($"""
                                                             <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                                               <w:p>
                                                                 <w:r>
                                                                   <w:t xml:space="preserve">Here comes a break</w:t>
                                                                 </w:r>
                                                               </w:p>
                                                               <w:p>
                                                                 <w:pPr>
                                                                   <w:sectPr>
                                                                     <w:type w:val="nextPage" />
                                                                   </w:sectPr>
                                                                 </w:pPr>
                                                               </w:p>
                                                               <w:p>
                                                                 <w:r>
                                                                   <w:t xml:space="preserve">and here comes the second part</w:t>
                                                                 </w:r>
                                                               </w:p>
                                                             </w:body>
                                                             """));
        }
    }
}
