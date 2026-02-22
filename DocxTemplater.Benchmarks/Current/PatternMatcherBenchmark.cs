using BenchmarkDotNet.Attributes;

namespace DocxTemplater.Benchmarks.Current
{
    [MemoryDiagnoser]
    public class PatternMatcherBenchmark
    {
        private const string SimpleText = "Hello {{Name}}!";
        private const string ComplexText = "Loop start {{#Items}} Item: {{Name}} - {{Price}:format(C2)} {{/Items}} End.";
        private const string ManyMatchesText = "{{a}}{{b}}{{c}}{{d}}{{e}}{{f}}{{g}}{{h}}{{i}}{{j}}{{k}}{{l}}{{m}}{{n}}{{o}}{{p}}";
        private const string NestedText = "{{#A}} {{#B}} {{C}} {{/B}} {{/A}}";

        [Benchmark]
        public void FindSyntaxPatternsSimple()
        {
            _ = PatternMatcher.FindSyntaxPatterns(SimpleText).ToList();
        }

        [Benchmark]
        public void FindSyntaxPatternsComplex()
        {
            _ = PatternMatcher.FindSyntaxPatterns(ComplexText).ToList();
        }

        [Benchmark]
        public void FindSyntaxPatternsManyMatches()
        {
            _ = PatternMatcher.FindSyntaxPatterns(ManyMatchesText).ToList();
        }

        [Benchmark]
        public void FindSyntaxPatternsNested()
        {
            _ = PatternMatcher.FindSyntaxPatterns(NestedText).ToList();
        }
    }
}
