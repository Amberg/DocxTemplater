using BenchmarkDotNet.Running;
using DocxTemplater.Benchmarks;

_ = BenchmarkRunner.Run<E2EBenchmark>();
