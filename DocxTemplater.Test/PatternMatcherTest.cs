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
            yield return new TestCaseData("{{/Items}}").Returns(new[] { PatternType.LoopEnd });
            yield return new TestCaseData("{{#items}}").Returns(new[] { PatternType.LoopStart });
            yield return new TestCaseData("{{/Items.InnerCollection}}").Returns(new[] { PatternType.LoopEnd });
            yield return new TestCaseData("{{#items.InnerCollection}}").Returns(new[] { PatternType.LoopStart });
            yield return new TestCaseData("{{a.foo > 5}}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{{ a > 5 }}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{{ a / 20 >= 12 }}").Returns(new[] { PatternType.Condition });
            yield return new TestCaseData("{{else}}").Returns(new[] { PatternType.ConditionElse });
            yield return new TestCaseData("{{/}}").Returns(new[] { PatternType.ConditionEnd });
            yield return new TestCaseData("NumericValue is greater than 0 - {{ds.Items.InnerCollection.InnerValue}:toupper()}{{else}}" +
                                          "I'm here if if this is not the case{{/}}{{/ds.Items.InnerCollection}}{{/Items}}")
                .Returns(new[] { PatternType.Variable, PatternType.ConditionElse, PatternType.ConditionEnd, PatternType.LoopEnd, PatternType.LoopEnd })
                .SetName("Complex Match 1");
        }
    }
}

