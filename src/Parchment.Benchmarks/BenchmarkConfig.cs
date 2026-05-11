class BenchmarkConfig : ManualConfig
{
    public BenchmarkConfig() =>
        AddJob(Job.Default.WithToolchain(InProcessEmitToolchain.Instance))
            .AddDiagnoser(MemoryDiagnoser.Default);
}
