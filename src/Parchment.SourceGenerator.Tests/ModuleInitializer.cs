public static partial class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        VerifySourceGenerators.Initialize();
        VerifierSettings.InitializePlugins();
        // The generator embeds template bytes as base64. The test docxes are rebuilt per run and
        // OPC packages are not byte-reproducible (zip timestamps, relationship ids), so the
        // literal is scrubbed — what the emission looks like is snapshotted, what the bytes were
        // is asserted by the convention tests within a single run.
        VerifierSettings.AddScrubber(builder =>
        {
            var scrubbed = Base64Literal().Replace(builder.ToString(), "FromBase64String(\"scrubbed\")");
            builder.Clear();
            builder.Append(scrubbed);
        });
    }

    [GeneratedRegex("""FromBase64String\("[A-Za-z0-9+/=]+"\)""")]
    private static partial Regex Base64Literal();
}
