using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Test
{
    /// <summary>
    /// https://github.com/Amberg/DocxTemplater/issues/128
    /// Hyperlinks, bookmarks and references break after generating.
    /// Bookmarks that are not affected by the templating (e.g. outside of loops)
    /// must be preserved so that cross-references (REF fields) and internal
    /// hyperlinks (w:anchor) keep working.
    /// </summary>
    internal class Issue128HyperlinksBookmarksReferences
    {
        [Test]
        public void BookmarkAndInternalHyperlinkArePreserved()
        {
            string content = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                                <w:bookmarkStart w:id=""1"" w:name=""targetHeading""/>
                                <w:r><w:t xml:space=""preserve"">Chapter {{ds}}</w:t></w:r>
                                <w:bookmarkEnd w:id=""1""/>
                              </w:p>
                              <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                                <w:hyperlink w:anchor=""targetHeading"">
                                  <w:r><w:t>see chapter above</w:t></w:r>
                                </w:hyperlink>
                              </w:p>";

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document
                {
                    Body = new Body { InnerXml = content }
                };
                wpDocument.Save();
            }
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", "hi there");
            var result = docTemplate.Process();
            docTemplate.Validate();

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;

            var bookmark = body.Descendants<BookmarkStart>().SingleOrDefault(x => x.Name == "targetHeading");
            Assert.That(bookmark, Is.Not.Null, "bookmark 'targetHeading' must be preserved so references keep working");
            Assert.That(body.Descendants<BookmarkEnd>().Count(x => x.Id == bookmark.Id), Is.EqualTo(1),
                "matching bookmarkEnd must be preserved");

            var hyperlink = body.Descendants<Hyperlink>().Single();
            Assert.That(hyperlink.Anchor?.Value, Is.EqualTo("targetHeading"),
                "internal hyperlink anchor must still point to an existing bookmark");
        }

        [Test]
        public void BookmarkInsideLoopIsClonedWithUniqueIdsAndNames()
        {
            string content = @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                                <w:r><w:t xml:space=""preserve"">{{#items}}</w:t></w:r>
                                <w:bookmarkStart w:id=""1"" w:name=""item""/>
                                <w:r><w:t xml:space=""preserve"">{{items}}</w:t></w:r>
                                <w:bookmarkEnd w:id=""1""/>
                                <w:r><w:t xml:space=""preserve"">{{/items}}</w:t></w:r>
                              </w:p>";

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document
                {
                    Body = new Body { InnerXml = content }
                };
                wpDocument.Save();
            }
            memStream.Position = 0;
            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("items", new[] { "a", "b", "c" });
            var result = docTemplate.Process();
            docTemplate.Validate(); // must produce a valid document

            using var document = WordprocessingDocument.Open(result, false);
            var body = document.MainDocumentPart.Document.Body;

            var starts = body.Descendants<BookmarkStart>().ToList();
            var ends = body.Descendants<BookmarkEnd>().ToList();
            Assert.Multiple(() =>
            {
                Assert.That(starts, Has.Count.EqualTo(3), "one bookmark per loop iteration");
                Assert.That(ends, Has.Count.EqualTo(3));
                // ids must be unique across the whole document
                Assert.That(starts.Select(x => x.Id.Value).Distinct().Count(), Is.EqualTo(3), "bookmark ids must be unique");
                // names must be unique too
                Assert.That(starts.Select(x => x.Name.Value).Distinct().Count(), Is.EqualTo(3), "bookmark names must be unique");
                // every start has a matching end with the same id
                Assert.That(starts.Select(x => x.Id.Value), Is.EquivalentTo(ends.Select(x => x.Id.Value)));
            });
        }
    }
}
