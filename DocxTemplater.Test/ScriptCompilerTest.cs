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
        public void CheckDirectComparisonOfPrimitiveTypes()
        {
            // default of type
            m_modelDictionary.Add("ds", 1.5m);
            Assert.That(m_scriptCompiler.CompileScript(". <= 2.0m")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(". > 1.0m")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(". > 1")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(". > -1")(), Is.True);
        }

        [Test]
        public void ConditionTestWithMultiplePrimitiveTypes()
        {
            var item = new { OpenAmount = -5.0m, IsCool = true, String = "myString", Count = 1 };

            m_modelDictionary.Add("Bills", new List<object> { item });

            var scope = m_modelDictionary.OpenScope();
            scope.AddVariable("item", item);

            Assert.That(m_scriptCompiler.CompileScript(".OpenAmount <= -5")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(".OpenAmount >= -4")(), Is.False);

            Assert.That(m_scriptCompiler.CompileScript(".IsCool == true")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(".IsCool == false")(), Is.False);

            Assert.That(m_scriptCompiler.CompileScript(".String == 'myString'")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(".String == 'yourString'")(), Is.False);

            Assert.That(m_scriptCompiler.CompileScript(".Count == 1")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript(".Count < 0")(), Is.False);
        }

        [Test]
        public void ConditionWithScopeAccessAndOperatorsWithoutSpaces()
        {
            m_modelDictionary.Add("x", new { a = new { b = 5, c = 25 } });
            var scope = m_modelDictionary.OpenScope();
            scope.AddVariable("foooo", 55);
            Assert.That(m_scriptCompiler.CompileScript(".>50&&.<60")());
            Assert.That(m_scriptCompiler.CompileScript(".%2==0")(), Is.False);

            var scope2 = m_modelDictionary.OpenScope();
            scope2.AddVariable("2", 550);
            Assert.That(m_scriptCompiler.CompileScript(".==..*10")());
        }

        [Test]
        public void WithMemberAccess()
        {
            m_modelDictionary.Add("x", new { a = new { b = 5 } });
            m_modelDictionary.Add("y", new
            {
                items = new[]
                {
                    new {b = 5},
                    new {b = 6}
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
        public void TestSpecialLoopVariables()
        {
            var model = new List<string[]>();
            model.Add(["Foo1", "Foo2", "test"]);
            model.Add(["Foo5", "Foo3", "test"]);

            m_modelDictionary.Add("Items", model);

            using var loopScope = m_modelDictionary.OpenScope();
            loopScope.AddVariable("Items", model[0]);
            loopScope.AddVariable($"Items._Idx", 0);
            loopScope.AddVariable($"Items._Length", model.Count);

            Assert.That(m_scriptCompiler.CompileScript("Items._Idx % 2 == 0")());
            Assert.That(m_scriptCompiler.CompileScript("Items._Length % 2 == 0")());
        }

        [Test]
        public void StringCompareAndFunctions()
        {
            m_modelDictionary.Add("x", new { a = new { b = "hi", c = "hi there", myDecimal = 5.0m } });
            m_modelDictionary.Add("y", new { a = "there" });
            m_modelDictionary.Add("c", "model is a string");
            m_modelDictionary.Add("d", 4);
            m_modelDictionary.Add("e", 4m); // decimal returns false for IsPrimitive


            Assert.That(m_scriptCompiler.CompileScript("x.a.b == \"hi\"")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b == 'by'")(), Is.False);
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('hi')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b == 'hi'")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('by')")(), Is.False);
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('hi')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.b.Contains('by')")(), Is.False);

            Assert.That(m_scriptCompiler.CompileScript("x.a.c.Contains(y.a)")());
            Assert.That(m_scriptCompiler.CompileScript("y.a.Contains(\"there\")")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.myDecimal > 2")()); // decimal returns false for IsPrimitive
            Assert.That(m_scriptCompiler.CompileScript("d % 2 == 0")());
            Assert.That(m_scriptCompiler.CompileScript("e % 2 == 0")());

        }
    }
}
