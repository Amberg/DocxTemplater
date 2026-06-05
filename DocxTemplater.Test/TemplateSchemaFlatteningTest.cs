using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Schema;

namespace DocxTemplater.Test
{
    /// <summary>
    /// Mirrors the variable-extraction scenarios a consumer builds on top of
    /// <see cref="DocxTemplate.GetTemplateSchema"/> (flatten the schema tree to a sorted,
    /// case-insensitively deduplicated list of dotted paths). The flatten helper here matches
    /// the semantics such a consumer needs:
    /// a node emits its own dotted path when it is a leaf (no properties), collections additionally
    /// emit their own path and recurse into their item schema.
    /// </summary>
    [TestFixture]
    internal sealed class TemplateSchemaFlatteningTest
    {
        private static IReadOnlyList<string> ExtractTokens(string body)
        {
            using var t = BuildTemplate(body);
            var schema = t.GetTemplateSchema();
            var tokens = new SortedSet<string>(StringComparer.Ordinal);
            foreach (var root in schema.Roots.Values)
            {
                Flatten(root.Name, root, tokens);
            }
            return tokens.ToList();
        }

        private static void Flatten(string path, TemplateSchemaNode node, ISet<string> acc)
        {
            if (node.Kind == TemplateNodeKind.Collection)
            {
                acc.Add(path);
                if (node.ItemSchema != null)
                {
                    foreach (var prop in node.ItemSchema.Properties.Values)
                    {
                        Flatten(path + "." + prop.Name, prop, acc);
                    }
                }
                return;
            }

            if (node.Properties.Count == 0)
            {
                acc.Add(path); // leaf
                return;
            }
            foreach (var prop in node.Properties.Values)
            {
                Flatten(path + "." + prop.Name, prop, acc);
            }
        }

        [Test]
        public void ReturnsScalarPaths()
        {
            Assert.That(ExtractTokens("Hello {{Customer}} — total: {{Total}}"),
                Is.EqualTo(["Customer", "Total"]));
        }

        [Test]
        public void FlattensDottedPaths()
        {
            Assert.That(ExtractTokens("{{Customer.Address.City}}"),
                Is.EqualTo(["Customer.Address.City"]));
        }

        [Test]
        public void FlattensLoopWithItemProperties()
        {
            Assert.That(ExtractTokens("{{#Rows}}{{Rows.Col}}{{/Rows}}"),
                Is.EqualTo(["Rows", "Rows.Col"]));
        }

        [Test]
        public void FlattensNestedObjects()
        {
            Assert.That(ExtractTokens("{{Customer.Address.City}} {{Customer.Name}}"),
                Is.EqualTo(["Customer.Address.City", "Customer.Name"]));
        }

        [Test]
        public void FlattensObjectInsideLoopItem()
        {
            Assert.That(ExtractTokens("{{#Rows}}{{Rows.Customer.Name}}{{/Rows}}"),
                Is.EqualTo(["Rows", "Rows.Customer.Name"]));
        }

        [Test]
        public void FlattensNestedCollections()
        {
            Assert.That(
                ExtractTokens("{{#Orders}}{{Orders.Id}}{{#Orders.Lines}}{{Orders.Lines.Sku}}{{/Orders.Lines}}{{/Orders}}"),
                Is.EqualTo(["Orders", "Orders.Id", "Orders.Lines", "Orders.Lines.Sku"]));
        }

        [Test]
        public void ExtractsPlainConditionVariable()
        {
            Assert.That(ExtractTokens("{?{Active}}yes{{/}}"), Is.EqualTo(["Active"]));
        }

        // Previously a known limitation (conditional *expressions* surfaced no variables); now fixed
        // via the SchemaExpressionParser detected-roots fallback. Member chains inside such an
        // expression degrade to their root (e.g. Customer, not Customer.IsHidden).
        [Test]
        public void ConditionWithNegation_YieldsRoot()
        {
            Assert.That(ExtractTokens("{?{!Active}}yes{{/}}"), Is.EqualTo(["Active"]));
        }

        [Test]
        public void ConditionWithComparison_YieldsRoot()
        {
            Assert.That(ExtractTokens("{?{Total > 0}}yes{{/}}"), Is.EqualTo(["Total"]));
        }

        [Test]
        public void ConditionWithBooleanLogic_YieldsRoots()
        {
            Assert.That(ExtractTokens("{?{ !Customer.IsHidden && Total > 0 }}yes{{/}}"),
                Is.EqualTo(["Customer", "Total"]));
        }

        [Test]
        public void DeduplicatesCaseInsensitively()
        {
            Assert.That(ExtractTokens("{{Name}} {{Name}} {{Other}}"),
                Is.EqualTo(["Name", "Other"]));
        }

        [Test]
        public void ReturnsEmpty_WhenNoTokens()
        {
            Assert.That(ExtractTokens("Just plain text — no placeholders."), Is.Empty);
        }

        private static DocxTemplate BuildTemplate(string body)
        {
            var stream = new MemoryStream();
            using (var wp = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = wp.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text(body)))));
                wp.Save();
            }
            stream.Position = 0;
            return new DocxTemplate(stream);
        }
    }
}
