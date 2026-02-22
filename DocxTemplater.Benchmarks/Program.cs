using BenchmarkDotNet.Running;
using DocxTemplater.Benchmarks;

#if BENCHMARK_CURRENT
// Run PatternMatcher (Current only)
_ = BenchmarkRunner.Run<PatternMatcherBenchmark>();
#endif

// Run Shared E2E
_ = BenchmarkRunner.Run<E2EBenchmark>();

// Run Shared Variable Evaluation
_ = BenchmarkRunner.Run<VariableEvaluationBenchmark>();
