using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTemplater.Test
{
    [TestFixture]
    public class SubTemplateImageImportTest
    {
        // Regression test: when a document is inserted via the 'template' formatter, images embedded in the
        // inserted document must be preserved. The body content is cloned into the target, so the referenced
        // image part has to be copied into the target and the blip relationship id remapped - otherwise the
        // r:embed dangles and the image is lost.
        [TestCase(true)]
        [TestCase(false)]
        public void SubTemplateInsert_PreservesEmbeddedImage(bool bindAsStream)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            // The document to insert: a placeholder (proves the deep merge still runs) and an embedded image.
            using var subDocStream = new MemoryStream();
            using (var subDocument = WordprocessingDocument.Create(subDocStream, WordprocessingDocumentType.Document))
            {
                var subMainPart = subDocument.AddMainDocumentPart();
                subMainPart.Document = new Document(new Body());
                var imagePart = subMainPart.AddImagePart(ImagePartType.Jpeg);
                using (var imageStream = new MemoryStream(imageBytes))
                {
                    imagePart.FeedData(imageStream);
                }
                var relationshipId = subMainPart.GetIdOfPart(imagePart);
                subMainPart.Document.Body.Append(
                    new Paragraph(new Run(new Text("Inserted for {{ds.Name}}"))),
                    new Paragraph(new Run(CreateInlineImage(relationshipId, 990000L, 792000L))));
                subMainPart.Document.Save();
            }

            using var memStream = new MemoryStream();
            using (var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wpDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("Start of Document"))),
                    new Paragraph(new Run(new Text("{{ds}:template('ds.SubDocument')}"))),
                    new Paragraph(new Run(new Text("End of Document")))));
            }
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", new
            {
                Name = "John",
                SubDocument = bindAsStream ? new MemoryStream(subDocStream.ToArray()) : (object)subDocStream.ToArray()
            });
            var result = docTemplate.Process();
            docTemplate.Validate();
            Assert.That(result, Is.Not.Null);
            result.Position = 0;

            using var document = WordprocessingDocument.Open(result, false);
            var mainDocumentPart = document.MainDocumentPart;
            var body = mainDocumentPart.Document.Body;

            // Placeholder inside the inserted document was resolved (deep merge still works).
            Assert.That(body.InnerText, Does.Contain("Inserted for John"));

            // The inserted image is present and its relationship resolves to an image part in the target.
            var blip = body.Descendants<A.Blip>().SingleOrDefault();
            Assert.That(blip, Is.Not.Null, "the inserted image drawing should be present in the merged document");
            Assert.That(blip.Embed?.Value, Is.Not.Null.And.Not.Empty);

            var referencedPart = mainDocumentPart.GetPartById(blip.Embed.Value);
            Assert.That(referencedPart, Is.InstanceOf<ImagePart>(), "the blip relationship must resolve to an image part in the target");

            // The copied image has the same content as the source image.
            using var copiedStream = ((ImagePart)referencedPart).GetStream();
            using var copiedBytes = new MemoryStream();
            copiedStream.CopyTo(copiedBytes);
            Assert.That(copiedBytes.ToArray(), Is.EqualTo(imageBytes), "the image content must be copied into the target document");
        }

        private static Drawing CreateInlineImage(string relationshipId, long cx, long cy)
        {
            var picture = new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = "image.jpg" },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

            var inline = new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.DocProperties { Id = 1U, Name = "Picture 1" },
                new A.Graphic(new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };
            return new Drawing(inline);
        }
    }
}
