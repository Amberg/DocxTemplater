using System.Collections;


namespace DocxTemplater.Test
{
    internal class PatternMatcherTest
    {

        [Test, TestCaseSource(nameof(TestPatternMatch_Cases))]
        public PatternType[] TestPatternMatch(string input)
        {
            var matches = PatternMatcher.FindSyntaxPatterns(input);
            foreach (var match in matches)
            {
                Assert.That(match.Match.Value.First(), Is.EqualTo('{'));
                Assert.That(match.Match.Value.Last(), Is.EqualTo('}'));
            }

            return matches.Select(x => x.Type).ToArray();
        }

        static IEnumerable TestPatternMatch_Cases()
        {
            yield return new TestCaseData("{{Foo}}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{Foo}:blupp()}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{/Items}}").Returns(new[] { PatternType.CollectionEnd });
            yield return new TestCaseData("{{ /Items }}").Returns(new[] { PatternType.CollectionEnd });
            yield return new TestCaseData("{{#items}}").Returns(new[] { PatternType.CollectionStart });
            yield return new TestCaseData("{{#.items}}").Returns(new[] { PatternType.CollectionStart });
            yield return new TestCaseData("{{#..items}}").Returns(new[] { PatternType.CollectionStart });
            yield return new TestCaseData("{{  #items  }}").Returns(new[] { PatternType.CollectionStart });
            yield return new TestCaseData("{{#ds.items_foo}}").Returns(new[] { PatternType.CollectionStart }).SetName("LoopStart Underscore dots");
            yield return new TestCaseData("{{/ds.items_foo}}").Returns(new[] { PatternType.CollectionEnd }).SetName("LoopEnd Underscore dots");
            yield return new TestCaseData("{{/Items.InnerCollection}}").Returns(new[] { PatternType.CollectionEnd });
            yield return new TestCaseData("{{#items.InnerCollection}}").Returns(new[] { PatternType.CollectionStart });
            yield return new TestCaseData("{?{ a.foo > 5}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{ a > 5 }}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{? { a > 5 } }").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{MyBool}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{!MyBool}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{ ? { MyBool}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{  a / 20 >= 12 }}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{  a / 20 >= 12 && a.b == true || !b.c }}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{{var}:F(d)}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{ds.foo.var}:f('HH : mm : s')}").Returns(new[] { PatternType.Variable }).SetName("Format with date pattern");
            yield return new TestCaseData("{{ds.foo.var}:f(HH:mm)}").Returns(new[] { PatternType.Variable }).SetName("Format with date pattern");
            yield return new TestCaseData("{{ds.foo.var}:f(0.##)}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{ds.foo.var}:F(d)}").Returns(new[] { PatternType.Variable }).SetName("Variable with dot");
            yield return new TestCaseData("{{ds.foo_blubb.var}:F(d)}").Returns(new[] { PatternType.Variable }).SetName("Variable with underscore");
            yield return new TestCaseData("{{var}:toupper}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{var}:f(MM/dd/yyy)}").Returns(new[] { PatternType.Variable }).SetName("Slash in pattern");
            yield return new TestCaseData("{{else}}").Returns(new[] { PatternType.ConditionElse });
            yield return new TestCaseData("{{  else  }}").Returns(new[] { PatternType.ConditionElse });
            yield return new TestCaseData("{{  :  }}").Returns(new[] { PatternType.ConditionElse });
            yield return new TestCaseData("{{:}}").Returns(new[] { PatternType.ConditionElse });
            yield return new TestCaseData("{{:s:}}").Returns(new[] { PatternType.CollectionSeparator });
            yield return new TestCaseData("{{: s :}}").Returns(new[] { PatternType.CollectionSeparator });
            yield return new TestCaseData("{{var}:format(a,b)}").Returns(new[] { PatternType.Variable }).SetName("Multiple Arguments");
            yield return new TestCaseData("{{/}}").Returns(new[] { PatternType.ConditionEnd });
            yield return new TestCaseData("{ { / } }").Returns(new[] { PatternType.ConditionEnd });
            yield return new TestCaseData("{?{ds.QrBills.idx == 2}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{?{ds.QrBills._Idx == 2}}").Returns(new[] { PatternType.Condition }).SetName("underscore in variable name");
            yield return new TestCaseData("{?{ds.QrBills._Idx % 2 == 0}}").Returns(new[] { PatternType.Condition }).SetName("modulo in condition");
            yield return new TestCaseData("{?{ds5.MyString.Contains('hi')}}").Returns(new[] { PatternType.Condition }).SetName("Text Contains");
            yield return new TestCaseData("{?{ds5 == 'hi'}}").Returns(new[] { PatternType.Condition }).SetName("Text Compare");
            yield return new TestCaseData("{?{x1 == \"hi\"}}").Returns(new[] { PatternType.Condition }).SetName("Text Compare Quoted with ");
            yield return new TestCaseData("{?{Ч1 == 'hi'}}").Returns(new[] { PatternType.Condition }).SetName("Text Compare Unicode");
            yield return new TestCaseData("{{Ч1}}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData("{{Ч1}:foo(Ч1)}").Returns(new[] { PatternType.Variable });
            yield return new TestCaseData(
                    "NumericValue is greater than 0 - {{ds.Items.InnerCollection.InnerValue}:toupper()}{{else}}" +
                    "I'm here if if this is not the case{{/}}{{/ds.Items.InnerCollection}}{{/Items}}")
                .Returns(new[]
                {
                    PatternType.Variable, PatternType.ConditionElse, PatternType.ConditionEnd, PatternType.CollectionEnd,
                    PatternType.CollectionEnd
                })
                .SetName("Complex Match 1");
        }

        [Test, TestCaseSource(nameof(PatternMatcherArgumentParsingTest_Cases))]
        public string[] PatternMatcherArgumentParsingTest(string syntax)
        {
            var match = PatternMatcher.FindSyntaxPatterns(syntax).First();
            return match.Arguments;
        }

        static IEnumerable PatternMatcherArgumentParsingTest_Cases()
        {
            yield return new TestCaseData("{{Foo}}").Returns(Array.Empty<string>());
            yield return new TestCaseData("{{Foo}:format}").Returns(Array.Empty<string>());
            yield return new TestCaseData("{{Foo}:format()}").Returns(Array.Empty<string>());
            yield return new TestCaseData("{{Foo}:format('MM/dd/yyyy')}").Returns(new[] { "MM/dd/yyyy" });
            yield return new TestCaseData("{{Foo}:format(MM/dd/yyyy)}").Returns(new[] { "MM/dd/yyyy" });
            yield return new TestCaseData("{{Foo}:format(\"MM/dd/yyyy\")}").Returns(new[] { "\"MM/dd/yyyy\"" });
            yield return new TestCaseData("{{Foo}:format('')}").Returns(new[] { string.Empty });
            yield return new TestCaseData("{{Foo}:format(a)}").Returns(new[] { "a" });
            yield return new TestCaseData("{{Foo}:format(param)}").Returns(new[] { "param" });
            yield return new TestCaseData("{{Foo}:format('param')}").Returns(new[] { "param" });
            yield return new TestCaseData("{{Foo}:format(a,b)}").Returns(new[] { "a", "b" });
            yield return new TestCaseData("{{Foo}:format(a,b,c)}").Returns(new[] { "a", "b", "c" });
            yield return new TestCaseData("{{Foo}:format(a.b, c.d)}").Returns(new[] { "a.b", "c.d" });
            yield return new TestCaseData("{{Foo}:format('a.b', c.d)}").Returns(new[] { "a.b", "c.d" });
            yield return new TestCaseData("{{Foo}:format(a,'a b',c)}").Returns(new[] { "a", "a b", "c" });
            yield return new TestCaseData("{{Foo}:format(a,b,'YYYY_MMM/DD FF',d)}").Returns(new[] { "a", "b", "YYYY_MMM/DD FF", "d" });
            yield return new TestCaseData("{{Foo}:format(a,'John Doe','YYYY_MMM/DD FF',d)}").Returns(new[] { "a", "John Doe", "YYYY_MMM/DD FF", "d" });
            yield return new TestCaseData("{{Foo}:f(a,'HH:mm',c)}").Returns(new[] { "a", "HH:mm", "c" });
            yield return new TestCaseData("{{Foo}:F(yyyy MM dd - HH mm ss)}").Returns(new[] { "yyyy MM dd - HH mm ss" });
            yield return new TestCaseData("{{Foo}:f(HH:mm,'HH:mm','HH : mm : ss')}").Returns(new[] { "HH:mm", "HH:mm", "HH : mm : ss" });
            yield return new TestCaseData("{{Foo}:f('comma in , argument', foo)}").Returns(new[] { "comma in , argument", "foo" });
            yield return new TestCaseData("{{Foo}:f(',', ',,')}").Returns(new[] { ",", ",," });
            yield return new TestCaseData("{{Foo}:f(' whitespacequoted ',   white space not quoted  )}").Returns(new[] { " whitespacequoted ", "white space not quoted  " });
            yield return new TestCaseData("{{Foo}:f('foo', 'this is \\'quoted\\' end')}").Returns(new[] { "foo", "this is 'quoted' end" });
            yield return new TestCaseData("{{Foo}:f('äöü', 'Foo \"blubb\" Test', \")}").Returns(new[] { "äöü", "Foo \"blubb\" Test", "\"" });
            yield return new TestCaseData("{{Foo}:img(w:14cm, h:15cm, r:90)}").Returns(new[] { "w:14cm", "h:15cm", "r:90" });

        }

        [TestCase("Some TExt {{    Variable   }} Some other text", ExpectedResult = "Variable")]
        [TestCase("Some TExt {{Variable   }} Some other text", ExpectedResult = "Variable")]
        [TestCase("Some TExt {{ Variable }} Some other text", ExpectedResult = "Variable")]
        public string AllowWhiteSpaceForVariables(string syntax)
        {
            var match = PatternMatcher.FindSyntaxPatterns(syntax).First();
            return match.Variable;
        }

        [TestCase("{?{Items._Idx % 2 == 0}}", ExpectedResult = "Items._Idx % 2 == 0")]
        [TestCase("{?{ds.QrBills._Idx % 2 == 0}}", ExpectedResult = "ds.QrBills._Idx % 2 == 0")]
        [TestCase("{?{ds.QrBills._Idx % 2 == 0 || ds.QrBills._Idx > 3 && x.y < 2}}", ExpectedResult = "ds.QrBills._Idx % 2 == 0 || ds.QrBills._Idx > 3 && x.y < 2")]
        [TestCase("{?{ds.QrBills._Idx}}", ExpectedResult = "ds.QrBills._Idx")]
        public string TestConditionExpression(string syntax)
        {
            var match = PatternMatcher.FindSyntaxPatterns(syntax).First();
            Assert.That(match.Type, Is.EqualTo(PatternType.Condition));
            return match.Condition;
        }
    }
}

