namespace DocxTemplater.Test
{
    internal class ModelLookupTest
    {
        [Test]
        public void LookupFirstModelDoesNotNeedAPrefix()
        {
            var modelLookup = new ModelLookup();
            modelLookup.Add("y", new { a = 6 });
            modelLookup.Add("x", new { a = 55 });
            modelLookup.Add("x.b", 42);
            Assert.That(modelLookup.GetValue("x.a"), Is.EqualTo(55));
            Assert.That(modelLookup.GetValue("x.b"), Is.EqualTo(42));
            Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(6));
            Assert.That(modelLookup.GetValue("a"), Is.EqualTo(6)); // y is added first.. y does not need a prefix
        }

        [Test]
        public void LookupFirstModelWithNestedPath()
        {
            var modelLookup = new ModelLookup();
            modelLookup.Add("y.a.b", new { c = 6 });
            Assert.That(modelLookup.GetValue("y.a.b.c"), Is.EqualTo(6));
            Assert.That(modelLookup.GetValue("a.b.c"), Is.EqualTo(6));

            modelLookup.Add("x.aa.bb", new { c = 55 });
            Assert.That(modelLookup.GetValue("x.aa.bb.c"), Is.EqualTo(55));
            Assert.Throws<OpenXmlTemplateException>(() => modelLookup.GetValue("aa.bb.c"));
        }

        [Test]
        public void ScopeVariablesAndImplicitAccessWithDot()
        {
            var modelLookup = new ModelLookup();
            modelLookup.Add("y", new { a = 6 });
            Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(6));

            using (var scope = modelLookup.OpenScope())
            {
                scope.AddVariable("y", new { a = 55 });
                Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(55));
            }
            Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(6));

            using (var outher = modelLookup.OpenScope())
            {
                outher.AddVariable("y", new { a = 66 });
                Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(66));
#pragma warning disable IDE0063 // Use simple 'using' statement
                using (var inner = modelLookup.OpenScope())
                {
                    inner.AddVariable("y", new { a = 77 });
                    Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(77));
                    Assert.That(modelLookup.GetValue(".a"), Is.EqualTo(77));
                    Assert.That(modelLookup.GetValue("..a"), Is.EqualTo(66));
                    Assert.That(modelLookup.GetValue("...a"), Is.EqualTo(6));
                }
#pragma warning restore IDE0063 // Use simple 'using' statement
            }

            // Add variable with leading dots to scope --> dots are removed
            using (var scope = modelLookup.OpenScope())
            {
                scope.AddVariable("...y", new { a = 55 });
                Assert.That(modelLookup.GetValue("y.a"), Is.EqualTo(55));
            }
        }
    }
}
