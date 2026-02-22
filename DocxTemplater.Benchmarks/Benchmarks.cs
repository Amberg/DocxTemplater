using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Dynamic;

namespace DocxTemplater.Benchmarks
{
    [MemoryDiagnoser]
    public class E2EBenchmark : IDisposable
    {
        private byte[] _templateBytes = null!;
        private object _model = null!;
        private DocxTemplate? _docTemplate;
        private MemoryStream? _ms;

        [GlobalSetup]
        public void Setup()
        {
            _templateBytes = CreateLargeTemplate(5000);
            _model = CreateLargeModel(5000);
        }

        private static byte[] CreateLargeTemplate(int count)
        {
            using var mem = new MemoryStream();
            using (var wordDocument = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body!;
                for (int i = 0; i < count; i++)
                {
                    var p = new Paragraph(new Run(new Text($"Property Number {i} is ")));
                    p.AppendChild(new Run(new Text($"{{{{ds.Prop{i}}}}}; ")));
                    body.AppendChild(p);
                }
            }
            return mem.ToArray();
        }

        private static object CreateLargeModel(int count)
        {
            var model = new ExpandoObject() as IDictionary<string, object>;
            for (int i = 0; i < count; i++)
            {
                model[$"Prop{i}"] = $"Value_{i}_of_many";
            }
            return model;
        }

        [IterationSetup]
        public void IterationSetup()
        {
            _ms = new MemoryStream(_templateBytes);
            _docTemplate = new DocxTemplate(_ms);
            _docTemplate.BindModel("ds", _model);
        }

        [IterationCleanup]
        public void IterationCleanup()
        {
            _docTemplate?.Dispose();
            _ms?.Dispose();
        }

        [Benchmark(Description = "Process Only (Placeholders)")]
        public void ProcessOnly()
        {
            _ = _docTemplate!.Process();
        }

        public void Dispose()
        {
            _docTemplate?.Dispose();
            _ms?.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    [MemoryDiagnoser]
    public class VariableEvaluationBenchmark : IDisposable
    {
        private byte[] _nestedTemplate = null!;
        private byte[] _formatterTemplate = null!;

        private object _nestedModel = null!;
        private object _simpleModel = null!;

        private DocxTemplate? _docTemplate;
        private MemoryStream? _ms;

        [GlobalSetup]
        public void Setup()
        {
            _nestedTemplate = CreateTemplate(1000, "{{ds.L1.L2.L3.L4.L5.Val}}");
            _nestedModel = new { L1 = new { L2 = new { L3 = new { L4 = new { L5 = new { Val = "Deep" } } } } } };
            _formatterTemplate = CreateTemplate(1000, "{{ds.Val}:toUpper()}");
            _simpleModel = new { Val = "text", BoolVal = true };
        }

        private static byte[] CreateTemplate(int count, string syntax)
        {
            using var mem = new MemoryStream();
            using (var wordDocument = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body!;
                for (int i = 0; i < count; i++)
                {
                    body.AppendChild(new Paragraph(new Run(new Text(syntax))));
                }
            }
            return mem.ToArray();
        }

        public void Dispose()
        {
            _docTemplate?.Dispose();
            _ms?.Dispose();
            GC.SuppressFinalize(this);
        }

        [Benchmark(Description = "Deep Nested Variables")]
        public void DeepNesting()
        {
            PrepareRun(_nestedTemplate, _nestedModel);
            _ = _docTemplate!.Process();
        }

        [Benchmark(Description = "Formatter Calls")]
        public void Formatters()
        {
            PrepareRun(_formatterTemplate, _simpleModel);
            _ = _docTemplate!.Process();
        }

        private void PrepareRun(byte[] template, object model)
        {
            _docTemplate?.Dispose();
            _ms?.Dispose();
            _ms = new MemoryStream(template);
            _docTemplate = new DocxTemplate(_ms);
            _docTemplate.BindModel("ds", model);
        }
    }

#if BENCHMARK_CURRENT
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
#endif
}
