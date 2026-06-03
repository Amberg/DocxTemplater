using DocxTemplater.Schema;

namespace DocxTemplater.Test
{
    [TestFixture]
    internal sealed class SchemaExpressionParserTest
    {
        private static List<TemplatePath> Collect(string expression)
        {
            var paths = new List<TemplatePath>();
            SchemaExpressionParser.Extract(expression, paths.Add);
            return paths;
        }

        [Test]
        public void SimpleMemberChain()
        {
            var paths = Collect("customer.Name");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].Root, Is.EqualTo("customer"));
            Assert.That(paths[0].Segments.Select(s => s.Name), Is.EqualTo(["Name"]));
        }

        [Test]
        public void NestedMemberChain()
        {
            var paths = Collect("customer.Address.City");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].Root, Is.EqualTo("customer"));
            Assert.That(paths[0].Segments.Select(s => s.Name), Is.EqualTo(["Address", "City"]));
        }

        [Test]
        public void TwoIndependentChains()
        {
            var paths = Collect("customer.Name + items.Count");
            Assert.That(paths.Select(p => p.ToString()), Is.EquivalentTo(["customer.Name", "items.Count"]));
        }

        [Test]
        public void BooleanShortCircuitDoesNotHideSecondOperand()
        {
            var paths = Collect("customer.IsPremium && customer.Score > 5");
            Assert.That(paths.Select(p => p.ToString()), Is.EquivalentTo(["customer.IsPremium", "customer.Score"]));
        }

        [Test]
        public void IndexerWithConstant()
        {
            var paths = Collect("items[0].Name");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].Root, Is.EqualTo("items"));
            Assert.That(paths[0].Segments.Select(s => s.Kind), Is.EqualTo([PathSegmentKind.Index, PathSegmentKind.Member]));
            Assert.That(paths[0].Segments[1].Name, Is.EqualTo("Name"));
        }

        [Test]
        public void MethodCall()
        {
            var paths = Collect("customer.Name.ToUpper()");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].Segments.Select(s => s.Kind), Is.EqualTo([PathSegmentKind.Member, PathSegmentKind.Method]));
            Assert.That(paths[0].Segments.Select(s => s.Name), Is.EqualTo(["Name", "ToUpper"]));
        }

        [Test]
        public void IndexerWithComputedArgumentDiscoversNestedChain()
        {
            // items[customer.idx].Name -> two roots
            var paths = Collect("items[customer.Idx].Name");
            var asStrings = paths.Select(p => p.ToString()).OrderBy(s => s).ToList();
            Assert.That(asStrings, Is.EqualTo(["customer.Idx", "items[].Name"]));
        }

        [Test]
        public void LeadingDot_ResolvesToParentScope()
        {
            // single dot = parent scope; customer is implicit
            var paths = Collect(".Name");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].LeadingDotCount, Is.EqualTo(1));
            Assert.That(paths[0].Root, Is.EqualTo("Name"));
            Assert.That(paths[0].Segments, Is.Empty);
        }

        [Test]
        public void LeadingDot_WithFurtherMembers()
        {
            var paths = Collect(".Address.City");
            Assert.That(paths, Has.Count.EqualTo(1));
            Assert.That(paths[0].LeadingDotCount, Is.EqualTo(1));
            Assert.That(paths[0].Root, Is.EqualTo("Address"));
            Assert.That(paths[0].Segments.Select(s => s.Name), Is.EqualTo(["City"]));
        }

        [Test]
        public void StringLiteralsAreIgnored()
        {
            var paths = Collect("customer.Name == \"items.Count\"");
            Assert.That(paths.Select(p => p.ToString()), Is.EquivalentTo(["customer.Name"]));
        }

        [Test]
        public void NumericLiteralsAreIgnored()
        {
            var paths = Collect("customer.Score > 5.5");
            Assert.That(paths.Select(p => p.ToString()), Is.EquivalentTo(["customer.Score"]));
        }

        [Test]
        public void EmptyExpression()
        {
            var paths = Collect("");
            Assert.That(paths, Is.Empty);
        }

        [Test]
        public void ExpressionWithOuterParens_TrailingToString_StaticallyResolved()
        {
            // (foo.bar.ToString()) - DynamicExpresso knows ToString() statically on `object`, so the
            // call becomes a regular MethodCallExpression (not a CallSite). The inner foo.bar is still
            // a CallSite chain and gets emitted. ToString does not appear in the schema (it's a
            // built-in, not a model property) - which is the correct outcome.
            var paths = Collect("(foo.bar.ToString())");
            Assert.That(paths.Select(p => p.ToString()), Is.EquivalentTo(["foo.bar"]));
        }

        // Tripwire tests for the DynamicExpresso internal binder shape. The collector reads
        // private fields and identifies binders by class name; a library upgrade that changes
        // either will silently lose paths unless these tests fail loudly.
        [Test]
        public void DynamicExpressoBinder_GetMember_StillReadable()
        {
            var paths = Collect("customer.Foo.Bar");
            Assert.That(paths, Has.Count.EqualTo(1), "TemplatePathCollector lost coupling to LateGetMemberCallSiteBinder");
            Assert.That(paths[0].Segments.Select(s => s.Name), Is.EqualTo(["Foo", "Bar"]));
        }

        [Test]
        public void DynamicExpressoBinder_InvokeMethod_StillReadable()
        {
            var paths = Collect("customer.DoSomething()");
            Assert.That(paths, Has.Count.EqualTo(1), "TemplatePathCollector lost coupling to LateInvokeMethodCallSiteBinder");
            Assert.That(paths[0].Segments, Has.Count.EqualTo(1));
            Assert.That(paths[0].Segments[0].Kind, Is.EqualTo(PathSegmentKind.Method));
            Assert.That(paths[0].Segments[0].Name, Is.EqualTo("DoSomething"));
        }

        [Test]
        public void DynamicExpressoBinder_GetIndex_StillReadable()
        {
            var paths = Collect("items[0]");
            Assert.That(paths, Has.Count.EqualTo(1), "TemplatePathCollector lost coupling to LateGetIndexCallSiteBinder");
            Assert.That(paths[0].Segments, Has.Count.EqualTo(1));
            Assert.That(paths[0].Segments[0].Kind, Is.EqualTo(PathSegmentKind.Index));
        }
    }
}
