#!/bin/bash
# Run benchmarks for Current and V26

echo "Running Current benchmarks..."
dotnet run -c Release --project DocxTemplater.Benchmarks/Current/DocxTemplater.Benchmarks.Current.csproj -- --job Quick --filter *

echo "--------------------------------------------------"
echo "Running V2.6 benchmarks..."
dotnet run -c Release --project DocxTemplater.Benchmarks/V26/DocxTemplater.Benchmarks.V26.csproj -- --job Quick --filter *

echo "--------------------------------------------------"
echo "Benchmarks finished."
echo "Results can be found in:"
echo "  Current: DocxTemplater.Benchmarks/Current/BenchmarkDotNet.Artifacts/results/"
echo "  V26:     DocxTemplater.Benchmarks/V26/BenchmarkDotNet.Artifacts/results/"
