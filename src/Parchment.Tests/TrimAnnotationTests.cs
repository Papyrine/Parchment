// Parchment reflects over the model at registration and constructs generic types at render and
// extraction, so NativeAOT is unsupported. These attributes are what tells a consumer that at
// publish time; without them the failure is a wrong exception or an empty table at runtime.
public class TrimAnnotationTests
{
    [Test]
    public async Task TemplateStoreEntryPointsWarnForTrimmingAndAot()
    {
        var methods = typeof(TemplateStore)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Where(_ => _.Name.StartsWith("Register") ||
                        _.Name.StartsWith("Render"))
            .ToList();

        await Assert.That(methods.Count).IsEqualTo(7);
        foreach (var method in methods)
        {
            await Assert.That(method.GetCustomAttribute<RequiresUnreferencedCodeAttribute>()).IsNotNull();
            await Assert.That(method.GetCustomAttribute<RequiresDynamicCodeAttribute>()).IsNotNull();
        }
    }

    [Test]
    public async Task ExtractWarnsForAot()
    {
        var methods = typeof(ParchmentExtractor)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Where(_ => _.Name == "Extract")
            .ToList();

        await Assert.That(methods.Count).IsEqualTo(2);
        foreach (var method in methods)
        {
            await Assert.That(method.GetCustomAttribute<RequiresDynamicCodeAttribute>()).IsNotNull();
        }
    }

    // IsTrimmable would let a trimmed publish trim the assembly without analysis warnings, which
    // is a claim of compatibility the library cannot make.
    [Test]
    public async Task AssemblyIsNotMarkedTrimmable()
    {
        var trimmable = typeof(TemplateStore).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .Any(_ => _.Key == "IsTrimmable");

        await Assert.That(trimmable).IsFalse();
    }
}
