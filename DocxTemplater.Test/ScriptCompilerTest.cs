namespace DocxTemplater.Test
{
    internal class ScriptCompilerTest
    {
        private ScriptCompiler m_scriptCompiler;
        private ModelLookup m_modelDictionary;

        [SetUp]
        public void Setup()
        {
            m_modelDictionary = new ModelLookup();
            m_scriptCompiler = new ScriptCompiler(m_modelDictionary, null);
        }

        [Test]
        public void ScriptWithoutMemberAccess()
        {
            Assert.That(m_scriptCompiler.CompileScript("10  / 2 == 5")());
            Assert.That(m_scriptCompiler.CompileScript("10  / 2 == 3")(), Is.False);
        }

        [Test]
        public void WithMemberAccess()
        {
            m_modelDictionary.Add("x", new { a = new { b = 5 } });
            m_modelDictionary.Add("y", new
            {
                items = new[]
                {
                    new { b = 5 },
                    new { b = 6 }
                }
            });
            var blockScope = m_modelDictionary.OpenScope();
            blockScope.AddVariable("y.items", new
            {
                b = 5
            });
            Assert.That(m_scriptCompiler.CompileScript("10  / 2 == x.a.b")());
            Assert.That(m_scriptCompiler.CompileScript("10  / 2 == y.items.b")());

        }

        [Test]
        public void StringCompareAndFunctions()
        {
            m_modelDictionary.Add("x", new { a = new { b = "hi", c = "hi there" } });
            m_modelDictionary.Add("y", new { a = "there" });
            Assert.That(m_scriptCompiler.CompileScript("x.a.b == \"hi\"")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b == 'by'")(), Is.False);
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('hi')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('by')")(), Is.False);
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('hi')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('by')")(), Is.False);

            Assert.That(m_scriptCompiler.CompileScript("x.a.c.Contains(y.a)")());

        }

    }

}
