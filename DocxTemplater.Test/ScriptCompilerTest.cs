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
            m_scriptCompiler = new ScriptCompiler(m_modelDictionary, new ProcessSettings { BindingErrorHandling = BindingErrorHandling.SkipBindingAndRemoveContent });
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

        [Test]
        public void ChainedStringOperations()
        {
            // Setup model.
            var a = new { b = new { c = "hi there" } };
            m_modelDictionary.Add("a", a);

            // Check parsing and execution of multiple chained string operations.
            Assert.That(m_scriptCompiler.CompileScript("a.b.c.Trim().EndsWith('there')")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript("a.b.c.Substring(2, 4).ToUpper().EndsWith(\"THE\")")(), Is.True);

            // Also check scope access / dot notation with multiple chained string operations.
            m_modelDictionary.OpenScope().AddVariable("a", a);
            Assert.That(m_scriptCompiler.CompileScript(".b.c.Substring(2).EndsWith('there')")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript("!.b.c.Substring(0, 2).EndsWith('hello')")(), Is.True);

            m_modelDictionary.OpenScope().AddVariable("b", a.b);
            Assert.That(m_scriptCompiler.CompileScript(".c.Substring(2).EndsWith('there')")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript("!.c.Substring(0, 2).EndsWith('hello')")(), Is.True);

            m_modelDictionary.OpenScope().AddVariable("c", a.b.c);
            Assert.That(m_scriptCompiler.CompileScript(".Substring(2).EndsWith('there')")(), Is.True);
            Assert.That(m_scriptCompiler.CompileScript("!.Substring(0, 2).EndsWith('hello')")(), Is.True);
        }
        [Test]
        public void CollectionMethodCall()
        {
            m_modelDictionary.Add("x", new { a = new { items = new List<string> { "hi", "there", "world" } } });

            Assert.That(m_scriptCompiler.CompileScript("x.a.items.Contains('hi')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.items.Contains('there')")());
            Assert.That(m_scriptCompiler.CompileScript("x.a.items.Contains('missing')")(), Is.False);
        }

        [Test]
        public void TestCompileExpression()
        {
            // Simple arithmetic
            Assert.That(m_scriptCompiler.CompileExpression("10 / 2")(), Is.EqualTo(5));

            // Null coalescing
            m_modelDictionary.Add("obj", new { Number = (int?)null, Text = "Hello" });
            Assert.That(m_scriptCompiler.CompileExpression("obj.Number ?? 5")(), Is.EqualTo(5));
            Assert.That(m_scriptCompiler.CompileExpression("obj.Number?.ToString() ?? \"N/A\"")(), Is.EqualTo("N/A"));
            Assert.That(m_scriptCompiler.CompileExpression("obj.Text ?? \"N/A\"")(), Is.EqualTo("Hello"));

            // Scope and dot notation
            m_modelDictionary.OpenScope().AddVariable("obj", new { Number = 15 });
            Assert.That(m_scriptCompiler.CompileExpression(".Number ?? 5")(), Is.EqualTo(15));
            Assert.That(m_scriptCompiler.CompileExpression(".Number?.ToString() ?? \"N/A\"")(), Is.EqualTo("15"));
        }

        [Test]
        public void ConditionWithMissingProperty_ReturnsNull()
        {
            m_modelDictionary.Add("ds", new { Name = "test" });

            // missing property in null check should resolve to null, not throw
            Assert.That(m_scriptCompiler.CompileScript("ds.MissingProp == null")());
        }

        [Test]
        public void ConditionWithStringIsNullOrWhiteSpace_OnMissingProperty()
        {
            m_modelDictionary.Add("ds", new { Name = "test" });

            // string.IsNullOrWhiteSpace on a missing property should return true
            Assert.That(m_scriptCompiler.CompileScript("string.IsNullOrWhiteSpace(ds.MissingProp)")());
            Assert.That(m_scriptCompiler.CompileScript("!string.IsNullOrWhiteSpace(ds.MissingProp)")(), Is.False);
        }

        [Test]
        public void ExpressionWithNullCoalescing_OnMissingProperty()
        {
            m_modelDictionary.Add("ds", new { Name = "test" });

            // null coalescing on a missing property should return the fallback
            Assert.That(m_scriptCompiler.CompileExpression("ds.MissingProp ?? \"fallback\"")(), Is.EqualTo("fallback"));
        }

        [Test]
        public void ExistingProperty_StillResolvesCorrectly()
        {
            m_modelDictionary.Add("ds", new { Name = "hello" });

            // existing properties must still work as before
            Assert.That(m_scriptCompiler.CompileScript("ds.Name == \"hello\"")());
            Assert.That(m_scriptCompiler.CompileScript("!string.IsNullOrWhiteSpace(ds.Name)")());
            Assert.That(m_scriptCompiler.CompileExpression("ds.Name ?? \"fallback\"")(), Is.EqualTo("hello"));
        }

        [Test]
        public void ConditionWithUnknownTopLevelIdentifier_ResolvesToNull()
        {
            m_modelDictionary.Add("ds", new { Name = "test" });

            // top-level identifier not in model (no ds. prefix) should resolve to null
            Assert.That(m_scriptCompiler.CompileScript("string.IsNullOrWhiteSpace(ClientBusinessType)")());
            Assert.That(m_scriptCompiler.CompileScript("!string.IsNullOrWhiteSpace(ClientBusinessType)")(), Is.False);
            Assert.That(m_scriptCompiler.CompileExpression("ClientBusinessType ?? \"fallback\"")(), Is.EqualTo("fallback"));
        }

        [Test]
        public void MethodInvocationOnDynamicObjects()
        {
            var model = new System.Collections.Hashtable
            {
                { "Nested", new { Name = "World" } },
                { "Value", 42 }
            };
            m_modelDictionary.Add("ds", model);

            // Method call on the dynamic object itself (Hashtable doesn't have much, but ToString works)
            Assert.That(m_scriptCompiler.CompileExpression("ds.ToString()")(), Is.Not.Null);

            // Method call on a property that is a "normal" object (resolved via DynamicObject/ModelVariable)
            Assert.That(m_scriptCompiler.CompileExpression("ds.Nested.ToString()")(), Is.EqualTo("{ Name = World }"));
            Assert.That(m_scriptCompiler.CompileExpression("ds.Nested.Name.ToUpper()")(), Is.EqualTo("WORLD"));

            // Method call on a simple type property
            Assert.That(m_scriptCompiler.CompileExpression("ds.Value.ToString()")(), Is.EqualTo("42"));
        }

        private class TestMethodModel
        {
            public string SayHello(string name)
            {
                return $"Hello {name}";
            }

            public string SayHello(string name, int count)
            {
                return $"Hello {name} {count}";
            }

            public string WithNull(string val)
            {
                return val ?? "was null";
            }

            public void Throw()
            {
                throw new InvalidOperationException("Test exception");
            }
        }

        [Test]
        public void MethodInvocation_WithComplexOverloadsAndNulls()
        {
            m_modelDictionary.Add("ds", new System.Collections.Hashtable { { "Model", new TestMethodModel() } });

            Assert.That(m_scriptCompiler.CompileExpression("ds.Model.SayHello('World')")(), Is.EqualTo("Hello World"));
            Assert.That(m_scriptCompiler.CompileExpression("ds.Model.SayHello('World', 42)")(), Is.EqualTo("Hello World 42"));
            Assert.That(m_scriptCompiler.CompileExpression("ds.Model.WithNull(null)")(), Is.EqualTo("was null"));
        }

        [Test]
        public void ToString_OnComplexProperty_ReturnsUnderlyingToString()
        {
            m_modelDictionary.Add("ds", new { Complex = new { Name = "RealObject" } });
            var expr = m_scriptCompiler.CompileExpression("ds.Complex");
            var result = expr();
            // result is a ScriptCompilerModelVariable, calling ToString() should return the underlying object's ToString
            Assert.That(result.ToString(), Is.EqualTo("{ Name = RealObject }"));
        }

        [Test]
        public void MethodInvocation_ThrowsException_WrappedInOpenXmlTemplateException()
        {
            m_modelDictionary.Add("ds", new System.Collections.Hashtable { { "Model", new TestMethodModel() } });
            var expr = m_scriptCompiler.CompileExpression("ds.Model.Throw()");
            var ex = Assert.Throws<OpenXmlTemplateException>(() => expr());
            Assert.That(ex.Message, Is.EqualTo("Test exception"));
        }

        [Test]
        public void BindingErrorHandling_ThrowException_ShouldThrowOnMissingProperty()
        {
            var settings = new ProcessSettings { BindingErrorHandling = BindingErrorHandling.ThrowException };
            var compiler = new ScriptCompiler(m_modelDictionary, settings);
            m_modelDictionary.Add("ds", new { Name = "Test" });

            var expr = compiler.CompileExpression("ds.MissingProp");
            Assert.Throws<OpenXmlTemplateException>(() => expr());
        }
    }
}
