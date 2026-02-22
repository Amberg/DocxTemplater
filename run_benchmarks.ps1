# Run benchmarks for Current and V26

Write-Host "Running Current benchmarks..." -ForegroundColor Cyan
dotnet run -c Release --project DocxTemplater.Benchmarks/Current/DocxTemplater.Benchmarks.Current.csproj -- --job Quick --filter *

Write-Host "--------------------------------------------------" -ForegroundColor Yellow
Write-Host "Running V2.6 benchmarks..." -ForegroundColor Cyan
dotnet run -c Release --project DocxTemplater.Benchmarks/V26/DocxTemplater.Benchmarks.V26.csproj -- --job Quick --filter *

Write-Host "--------------------------------------------------" -ForegroundColor Yellow
Write-Host "Benchmarks finished." -ForegroundColor Green
Write-Host "Results can be found in:"
Write-Host "  Current: DocxTemplater.Benchmarks/Current/BenchmarkDotNet.Artifacts/results/"
Write-Host "  V26:     DocxTemplater.Benchmarks/V26/BenchmarkDotNet.Artifacts/results/"
