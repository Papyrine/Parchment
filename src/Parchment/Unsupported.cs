/// <summary>
/// The messages a consumer sees from the trim and AOT analyzers when calling into the parts of
/// Parchment that reflect. See "Trimming and NativeAOT" in contributing.md for the sites.
/// </summary>
static class Unsupported
{
    public const string Trimming =
        "Template tokens are validated against the model through reflection, and Excelsior reflects over the element type of an [ExcelsiorTable] collection. Trimming can remove the members both look for.";

    public const string DynamicCode =
        "Rendering an [ExcelsiorTable] member and extracting an editable collection construct generic types and methods at runtime, for which NativeAOT has no native code. Parchment does not support NativeAOT.";
}
