using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Model;
using System.ComponentModel;

namespace DocxTemplater.Test
{
    /// <summary>
    /// https://github.com/Amberg/DocxTemplater/issues/167 - public fields are bound like properties.
    /// </summary>
    internal class Issue167PublicFieldsTest
    {
        private class ModelWithFields
        {
            public int Field = 1;
            public int Property { get; set; } = 2;
            public ModelWithFields Nested;

            public string Both = "field";
            public string BOTH { get; set; } = "property";

            public static int StaticField = 99;

            [ModelProperty(DefaultFormatter = "toupper")]
            public string Shout = "quiet";
        }

        private class DisplayNameModel : TemplateModelWithDisplayNames
        {
            [DisplayName("Vorname")]
            public string FirstName { get; set; } = "Ada";

            public string LastName = "Lovelace";
        }

        [Test]
        public void FieldsAreResolvedLikeProperties()
        {
            var lookup = new ModelLookup();
            lookup.Add("ds", new ModelWithFields { Nested = new ModelWithFields { Field = 3 } });
            Assert.Multiple(() =>
            {
                Assert.That(lookup.GetValue("ds.Field"), Is.EqualTo(1));
                Assert.That(lookup.GetValue("ds.Property"), Is.EqualTo(2));
                Assert.That(lookup.GetValue("ds.FIELD"), Is.EqualTo(1), "case-insensitive like properties");
                Assert.That(lookup.GetValue("ds.Nested.Field"), Is.EqualTo(3));
                Assert.That(lookup.GetValue("ds.Nested.Nested.Field"), Is.Null, "null propagation through a null field");
            });
        }

        [Test]
        public void PropertyTakesPrecedenceOverFieldWithTheSameName()
        {
            var lookup = new ModelLookup();
            lookup.Add("ds", new ModelWithFields());
            Assert.That(lookup.GetValue("ds.Both"), Is.EqualTo("property"));
        }

        [Test]
        public void StaticFieldsAreNotBound()
        {
            var lookup = new ModelLookup();
            lookup.Add("ds", new ModelWithFields());
            var ex = Assert.Throws<OpenXmlTemplateException>(() => lookup.GetValue("ds.StaticField"));
            Assert.That(ex.ErrorCode, Is.EqualTo(TemplateErrorCode.PropertyNotFoundOnType));
        }

        [Test]
        public void DefaultFormatterAttributeOnField()
        {
            var lookup = new ModelLookup();
            lookup.Add("ds", new ModelWithFields());
            Assert.That(lookup.GetValueWithMetadata("ds.Shout").Metadata.DefaultFormatter, Is.EqualTo("toupper"));
        }

        [Test]
        public void ValueTupleItemsAreFields()
        {
            var lookup = new ModelLookup();
            // the element names of a named tuple do not exist at runtime - the items are the fields Item1, Item2, ...
            lookup.Add("item", (Name: "a", Price: 1));
            Assert.Multiple(() =>
            {
                Assert.That(lookup.GetValue("item.Item1"), Is.EqualTo("a"));
                Assert.That(lookup.GetValue("item.Item2"), Is.EqualTo(1));
            });
        }

        [Test]
        public void TemplateModelWithDisplayNamesResolvesFieldsByName()
        {
            var lookup = new ModelLookup();
            lookup.Add("ds", new DisplayNameModel());
            Assert.Multiple(() =>
            {
                Assert.That(lookup.GetValue("ds.Vorname"), Is.EqualTo("Ada"));
                Assert.That(lookup.GetValue("ds.lastname"), Is.EqualTo("Lovelace"));
            });
        }

        [Test]
        public void ListOfValueTuplesAndFieldWithDefaultFormatterInDocument()
        {
            var model = new
            {
                Shout = new ModelWithFields().Shout,
                Fields = new ModelWithFields(),
                Items = new List<(string Name, int Price)> { ("a", 1), ("b", 2) }
            };
            var text = Render("{{ds.Fields.Shout}} {{#ds.Items}}{{.Item1}}={{.Item2}};{{/ds.Items}}", model);
            Assert.That(text, Is.EqualTo("QUIET a=1;b=2;"));
        }

        private static string Render(string template, object model)
        {
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(template)))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.BindModel("ds", model);
            var result = docTemplate.Process();
            docTemplate.Validate();
            var document = WordprocessingDocument.Open(result, false);
            return document.MainDocumentPart.Document.Body.InnerText;
        }
    }
}
