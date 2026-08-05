using Parchment;

// One model per template, so a template the generator cannot see shows up as PARCH004 rather than
// as silence.

[ParchmentModel("Templates/copied.md")]
public partial class CopiedModel
{
    public required string Name { get; init; }
}

[ParchmentModel("Templates/embedded.md")]
public partial class EmbeddedModel
{
    public required string Name { get; init; }
}

[ParchmentModel("Templates/dual.md")]
public partial class DualModel
{
    public required string Name { get; init; }
}
