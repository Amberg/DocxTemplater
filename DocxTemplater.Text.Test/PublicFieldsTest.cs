using System.Globalization;

namespace DocxTemplater.Text.Test
{
    /// <summary>
    /// https://github.com/Amberg/DocxTemplater/issues/167 - public fields are bound like properties, in text templates too.
    /// </summary>
    internal class PublicFieldsTest
    {
        [Test]
        public void ListOfValueTuples()
        {
            var settings = new TextTemplateSettings { Escaping = TextEscaping.None, Culture = CultureInfo.InvariantCulture };
            var template = new TextTemplate("{{#ds.Items}}{{.Item1}}={{.Item2}};{{/ds.Items}}", settings);
            template.BindModel("ds", new { Items = new List<(string Name, int Price)> { ("a", 1), ("b", 2) } });
            Assert.That(template.Process().Text, Is.EqualTo("a=1;b=2;"));
        }
    }
}
