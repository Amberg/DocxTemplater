using DocxTemplater.Schema;

namespace DocxTemplater.Test
{
    [TestFixture]
    internal sealed class SchemaBuilderTest
    {
        [Test]
        public void DeclareScalar_AtRoot()
        {
            var b = new SchemaBuilder();
            b.DeclareScalar("customer.Name");
            var schema = b.Build();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["customer"]));
            var customer = schema.Roots["customer"];
            Assert.That(customer.Kind, Is.EqualTo(TemplateNodeKind.Object));
            Assert.That(customer.Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void DeclareCollection_AtRoot()
        {
            var b = new SchemaBuilder();
            var item = b.DeclareCollection("items");
            b.OpenItemScope("items", item);
            b.DeclareScalar("items.Price");
            b.DeclareScalar("items.Name");
            b.CloseScope();
            var schema = b.Build();

            var items = schema.Roots["items"];
            Assert.That(items.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(items.ItemSchema, Is.Not.Null);
            Assert.That(items.ItemSchema.Properties.Keys, Is.EquivalentTo(["Price", "Name"]));
        }

        [Test]
        public void NestedCollection_DottedScopeKey()
        {
            var b = new SchemaBuilder();
            var outerItem = b.DeclareCollection("customer.orders");
            b.OpenItemScope("customer.orders", outerItem);
            b.DeclareScalar("customer.orders.amount");
            b.CloseScope();
            var schema = b.Build();

            var customer = schema.Roots["customer"];
            var orders = customer.Properties["orders"];
            Assert.That(orders.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(orders.ItemSchema.Properties["amount"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void OuterModelStillResolvesInsideLoop()
        {
            // Inside a loop on `items`, paths starting with `customer` (not the loop's name) must fall through to root.
            var b = new SchemaBuilder();
            var item = b.DeclareCollection("items");
            b.OpenItemScope("items", item);
            b.DeclareScalar("items.Price");
            b.DeclareScalar("customer.Name");
            b.CloseScope();
            var schema = b.Build();

            Assert.That(schema.Roots.Keys, Is.EquivalentTo(["items", "customer"]));
            Assert.That(schema.Roots["customer"].Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void LeadingDot_ResolvesToInnerScope()
        {
            var b = new SchemaBuilder();
            var item = b.DeclareCollection("items");
            b.OpenItemScope("items", item);
            // `.Price` inside loop → property of the current item
            b.DeclareScalar(new TemplatePath("Price", []) { LeadingDotCount = 1 });
            b.CloseScope();
            var schema = b.Build();

            Assert.That(schema.Roots["items"].ItemSchema.Properties["Price"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void LeadingDoubleDot_SkipsOneScope()
        {
            var b = new SchemaBuilder();
            var outerItem = b.DeclareCollection("orders");
            b.OpenItemScope("orders", outerItem);
            var innerItem = b.DeclareCollection("orders.lines");
            b.OpenItemScope("orders.lines", innerItem);
            // `..Customer` from inner loop = one level up = orders item
            b.DeclareScalar(new TemplatePath("Customer", []) { LeadingDotCount = 2 });
            b.CloseScope();
            b.CloseScope();
            var schema = b.Build();

            Assert.That(schema.Roots["orders"].ItemSchema.Properties.Keys, Does.Contain("Customer"));
        }

        [Test]
        public void LoopIndexAndLengthAreIgnored()
        {
            var b = new SchemaBuilder();
            var item = b.DeclareCollection("items");
            b.OpenItemScope("items", item);
            b.DeclareScalar("items._Idx");
            b.DeclareScalar("items._Length");
            b.CloseScope();
            var schema = b.Build();

            Assert.That(schema.Roots["items"].ItemSchema.Properties, Is.Empty);
        }

        [Test]
        public void DeclarePath_IndexerPromotesToCollection()
        {
            // items[0].Name -> items becomes a Collection whose ItemSchema has the Name property
            var b = new SchemaBuilder();
            var path = new TemplatePath("items",
            [
                new PathSegment { Kind = PathSegmentKind.Index },
                new PathSegment { Kind = PathSegmentKind.Member, Name = "Name" }
            ]);
            b.DeclareScalar(path);
            var schema = b.Build();

            var items = schema.Roots["items"];
            Assert.That(items.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(items.ItemSchema.Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }

        [Test]
        public void DeclarePath_MethodCallDropsSubsequentSegments()
        {
            // customer.Name.ToUpper() -> ToUpper is a method call, not a property; subsequent paths from
            // the method's return value are unknown and should not appear in the schema.
            var b = new SchemaBuilder();
            var path = new TemplatePath("customer",
            [
                new PathSegment { Kind = PathSegmentKind.Member, Name = "Name" },
                new PathSegment { Kind = PathSegmentKind.Method, Name = "ToUpper" }
            ]);
            b.DeclareScalar(path);
            var schema = b.Build();

            var customer = schema.Roots["customer"];
            Assert.That(customer.Properties.Keys, Is.EquivalentTo(["Name"]));
            Assert.That(customer.Properties["Name"].Properties, Is.Empty, "method name should not appear as a property");
        }

        [Test]
        public void DottedCollectionWithLeadingDotInsideOuterLoop()
        {
            // Loop A on `outer`, inside it loop B on `.inner` (a property of the outer item).
            // We need to declare the collection at the outer item's schema.
            var b = new SchemaBuilder();
            var outerItem = b.DeclareCollection("outer");
            b.OpenItemScope("outer", outerItem);
            // Inside the outer loop, declare a collection called `.inner`
            var innerItem = b.DeclareCollection(".inner");
            b.OpenItemScope("outer.inner", innerItem); // analyzer reconstructs the absolute key
            b.DeclareScalar(new TemplatePath("Name", []) { LeadingDotCount = 1 });
            b.CloseScope();
            b.CloseScope();
            var schema = b.Build();

            var outerNode = schema.Roots["outer"];
            Assert.That(outerNode.ItemSchema.Properties.Keys, Does.Contain("inner"));
            var inner = outerNode.ItemSchema.Properties["inner"];
            Assert.That(inner.Kind, Is.EqualTo(TemplateNodeKind.Collection));
            Assert.That(inner.ItemSchema.Properties["Name"].Kind, Is.EqualTo(TemplateNodeKind.Scalar));
        }
    }
}
