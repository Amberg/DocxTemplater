using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Schema;

namespace DocxTemplater.Test
{
    [TestFixture]
    internal sealed class TemplateSchemaTest
    {
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

        [Test]
        public void SimpleVariable()
        {
            using var t = BuildTemplate("Hello {{customer.Name}}!");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["customer"]));
            Assert.That(schema.Roots["customer"].Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void Loop_ProducesCollectionWithItemSchema()
        {
            using var t = BuildTemplate("{{#items}}{{items.Price}}{{/items}}");
            var schema = t.GetTemplateSchema();

            var items = schema.Roots["items"];
            Assert.That(items.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(items.ItemSchema.Properties["Price"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void Loop_ReferencingOuterModel()
        {
            using var t = BuildTemplate("{{#items}}{{items.Price}} for {{customer.Name}}{{/items}}");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["items", "customer"]));
            Assert.That(schema.Roots["items"].ItemSchema.Properties["Price"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
            Assert.That(schema.Roots["customer"].Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void NestedLoop_DottedScopeKey()
        {
            using var t = BuildTemplate("{{#customer.orders}}{{customer.orders.amount}}{{/customer.orders}}");
            var schema = t.GetTemplateSchema();

            var customer = schema.Roots["customer"];
            var orders = customer.Properties["orders"];
            Assert.That(orders.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(orders.ItemSchema.Properties["amount"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void NestedLoopWithLeadingDot()
        {
            using var t = BuildTemplate("{{#orders}}{{#.lines}}{{.qty}}{{/.lines}}{{/orders}}");
            var schema = t.GetTemplateSchema();

            var orders = schema.Roots["orders"];
            Assert.That(orders.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            var lines = orders.ItemSchema.Properties["lines"];
            Assert.That(lines.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(lines.ItemSchema.Properties["qty"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void Condition_BothBranchesContributeToSchema()
        {
            using var t = BuildTemplate("{?{customer.IsPremium}}{{customer.PremiumLevel}}{{else}}{{customer.StandardLevel}}{{/}}");
            var schema = t.GetTemplateSchema();

            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Is.EquivalentTo(["IsPremium", "PremiumLevel", "StandardLevel"]));
        }

        [Test]
        public void ConditionExpression_CollectsMemberPaths()
        {
            using var t = BuildTemplate("{?{customer.IsPremium && customer.Score > 5}}yes{{/}}");
            var schema = t.GetTemplateSchema();

            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Is.EquivalentTo(["IsPremium", "Score"]));
        }

        [Test]
        public void ExpressionBlock_CollectsVariables()
        {
            using var t = BuildTemplate("Total: {{(items.Count * customer.Multiplier)}}");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots["items"].Properties.Keys, Does.Contain("Count"));
            Assert.That(schema.Roots["customer"].Properties.Keys, Does.Contain("Multiplier"));
        }

        [Test]
        public void Switch_DeclaresSwitchVariable()
        {
            using var t = BuildTemplate("{{#switch:customer.Tier}}{{#case:1}}A{{/}}{{#case:2}}B{{/}}{{/}}");
            var schema = t.GetTemplateSchema();

            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Does.Contain("Tier"));
        }

        [Test]
        public void IndexerInExpression_PromotesToCollection()
        {
            using var t = BuildTemplate("{{(items[0].Name)}}");
            var schema = t.GetTemplateSchema();

            var items = schema.Roots["items"];
            Assert.That(items.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(items.ItemSchema.Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void MethodCallDoesNotLeakAsProperty()
        {
            using var t = BuildTemplate("{{(customer.Name.ToUpper())}}");
            var schema = t.GetTemplateSchema();

            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Is.EquivalentTo(["Name"]));
            Assert.That(customer.Properties["Name"].Properties, Is.Empty);
        }

        [Test]
        public void LoopIndexAndLengthAreFiltered()
        {
            using var t = BuildTemplate("{{#items}}{{items._Idx}}/{{items._Length}}{{/items}}");
            var schema = t.GetTemplateSchema();

            var items = schema.Roots["items"];
            // _Idx and _Length are loop-internal — schema should not require them
            Assert.That(items.ItemSchema.Properties.Keys, Does.Not.Contain("_Idx"));
            Assert.That(items.ItemSchema.Properties.Keys, Does.Not.Contain("_Length"));
        }

        [Test]
        public void IgnoredContent_IsNotIncluded()
        {
            using var t = BuildTemplate("{{customer.Name}}{{:ignore}}{{secret.ApiKey}}{{/:ignore}}");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["customer"]));
        }

        [Test]
        public void EmptyTemplate_YieldsEmptySchema()
        {
            using var t = BuildTemplate("Just some static text.");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots, Is.Empty);
        }

        [Test]
        public void Schema_CombinedFeatures()
        {
            // A realistic template that exercises loops, nested loops, conditions, expressions and
            // an outer-model reference, all in one go.
            using var t = BuildTemplate(
                "Hello {{customer.Name}} "
                + "{?{customer.IsPremium}}Premium since {{customer.JoinedAt}}{{/}}"
                + "{{#customer.orders}}"
                + "Order {{customer.orders.Id}}: total {{(customer.orders.Total)}}"
                + "{{#.lines}}{{.Product}} x {{.Qty}}{{/.lines}}"
                + "{{/customer.orders}}");
            var schema = t.GetTemplateSchema();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["customer"]));
            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Is.EquivalentTo(["Name", "IsPremium", "JoinedAt", "orders"]));

            var orders = customer.Properties["orders"];
            Assert.That(orders.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(orders.ItemSchema.Properties.Keys, Is.EquivalentTo(["Id", "Total", "lines"]));

            var lines = orders.ItemSchema.Properties["lines"];
            Assert.That(lines.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(lines.ItemSchema.Properties.Keys, Is.EquivalentTo(["Product", "Qty"]));
        }
    }
}
