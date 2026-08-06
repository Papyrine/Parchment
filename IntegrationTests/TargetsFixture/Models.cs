using Parchment;

// One model per template, so a template the generator cannot see shows up as PARCH004 rather than
// as silence.

[ParchmentModel]
public partial class CopiedModel
{
    public required string Name { get; init; }
}

[ParchmentModel]
public partial class DualModel
{
    public required string Name { get; init; }
}
