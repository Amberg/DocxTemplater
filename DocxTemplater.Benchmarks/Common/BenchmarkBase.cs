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
            // We do NOT dispose docTemplate in IterationCleanup to avoid closing the stream too early for BenchmarkDotNet
            _docTemplate = new DocxTemplate(_ms);
            _docTemplate.BindModel("ds", _model);
        }

        [Benchmark(Description = "Process Only (5000 Placeholders)")]
        public void ProcessOnly()
        {
            // Just call process, do not dispose the result stream here
            _ = _docTemplate!.Process();
        }

        public void Dispose()
        {
            _docTemplate?.Dispose();
            _ms?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
