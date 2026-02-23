using System.Diagnostics;
using System.Text;

namespace DocxTemplater.Benchmarks.Runner
{
    internal sealed class Program
    {
        private static readonly string ArtifactsRoot = Path.Combine(Environment.CurrentDirectory, "DocxTemplater.Benchmarks.Runner", "artifacts");
        private const string ReportFile = "BENCHMARK_REPORT.md";
        private static readonly string[] Versions = ["Current", "2.6.1"];
        private static readonly string[] Frameworks = ["net8.0", "net9.0", "net10.0"];

        private static async Task Main()
        {
            Console.WriteLine("=== DocxTemplater Benchmark Orchestrator ===");

            if (Directory.Exists(ArtifactsRoot))
            {
                TryDeleteDirectory(ArtifactsRoot);
            }
            Directory.CreateDirectory(ArtifactsRoot);

            foreach (var framework in Frameworks)
            {
                foreach (var version in Versions)
                {
                    RunBenchmarks(version, framework);
                }
            }

            await GenerateConsolidatedReport();

            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"Benchmarks finished. Summary available in {ReportFile}");
        }

        private static void RunBenchmarks(string version, string framework)
        {
            Console.WriteLine($"\n>>> Running benchmarks for version: {version} on {framework}...");

            string projectFile = Path.Combine("DocxTemplater.Benchmarks", "DocxTemplater.Benchmarks.csproj");

            if (Directory.Exists("BenchmarkDotNet.Artifacts"))
            {
                TryDeleteDirectory("BenchmarkDotNet.Artifacts");
            }

            var processStartInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"run -c Release --project \"{projectFile}\" -f {framework} -p:BenchmarkVersion={version} -- --job Quick --filter *",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Environment.CurrentDirectory
            };

            using var process = Process.Start(processStartInfo) ?? throw new InvalidOperationException("Failed to start dotnet run");

            process.OutputDataReceived += (s, e) =>
            {
                if (e.Data != null)
                {
                    Console.WriteLine(e.Data);
                }
            };
            process.ErrorDataReceived += (s, e) =>
            {
                if (e.Data != null)
                {
                    Console.Error.WriteLine(e.Data);
                }
            };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                Console.WriteLine($"Benchmarks for {version} on {framework} failed with exit code {process.ExitCode}");
                return;
            }

            var targetDir = Path.Combine(ArtifactsRoot, $"{version}_{framework}");
            if (Directory.Exists("BenchmarkDotNet.Artifacts"))
            {
                TryMoveDirectory("BenchmarkDotNet.Artifacts", targetDir);
            }
        }

        private static void TryDeleteDirectory(string path)
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    Directory.Delete(path, true);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(1000);
                }
            }
        }

        private static void TryMoveDirectory(string source, string dest)
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    if (Directory.Exists(dest))
                    {
                        Directory.Delete(dest, true);
                    }
                    Directory.Move(source, dest);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(1000);
                }
            }
            Console.WriteLine($"Failed to move {source} to {dest} after retries.");
        }

        private static async Task GenerateConsolidatedReport()
        {
            Console.WriteLine("\n>>> Consolidating results...");
            var sb = new StringBuilder();
            sb.AppendLine("# Performance Comparison Report");
            sb.AppendLine($"Generated on: {DateTime.Now:F}");
            sb.AppendLine();
            sb.AppendLine("## End-to-End & Variable Evaluation Results");
            sb.AppendLine();
            sb.AppendLine("| Version | Framework | Method | Mean | Allocated |");
            sb.AppendLine("|--- |--- |--- |--- |--- |");

            foreach (var framework in Frameworks)
            {
                foreach (var version in Versions)
                {
                    var resultsPath = Path.Combine(ArtifactsRoot, $"{version}_{framework}", "results");
                    if (!Directory.Exists(resultsPath))
                    {
                        continue;
                    }

                    foreach (var file in Directory.GetFiles(resultsPath, "*-report-github.md"))
                    {
                        if (file.Contains("PatternMatcher"))
                        {
                            continue;
                        }

                        var lines = await File.ReadAllLinesAsync(file);
                        foreach (var line in lines)
                        {
                            var trimmedLine = line.Trim();
                            // Match lines starting with | and containing method names (encoded or not)
                            if (trimmedLine.StartsWith("| &#39;") || trimmedLine.StartsWith("| '"))
                            {
                                var cleanLine = trimmedLine.Replace("&#39;", "'");
                                sb.AppendLine($"| {version} | {framework} {cleanLine}");
                            }
                        }
                    }
                }
            }

            sb.AppendLine();
            sb.AppendLine("## Pattern Matcher (Current only)");
            sb.AppendLine();

            foreach (var framework in Frameworks)
            {
                sb.AppendLine($"### Framework: {framework}");
                sb.AppendLine();
                var resultsPath = Path.Combine(ArtifactsRoot, $"Current_{framework}", "results");
                if (!Directory.Exists(resultsPath))
                {
                    continue;
                }

                var matcherReport = Directory.GetFiles(resultsPath, "*PatternMatcherBenchmark-report-github.md").FirstOrDefault();
                if (matcherReport != null)
                {
                    var lines = await File.ReadAllLinesAsync(matcherReport);
                    bool inTable = false;
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("| Method"))
                        {
                            inTable = true;
                        }
                        if (inTable)
                        {
                            sb.AppendLine(line);
                            if (string.IsNullOrWhiteSpace(line))
                            {
                                inTable = false;
                            }
                        }
                    }
                }
                sb.AppendLine();
            }

            string reportPath = Path.Combine(Environment.CurrentDirectory, ReportFile);
            await File.WriteAllTextAsync(reportPath, sb.ToString());
        }
    }
}
