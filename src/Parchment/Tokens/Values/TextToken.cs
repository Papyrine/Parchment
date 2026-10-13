/// <summary>
/// What the implicit conversion from <c>string</c> on <see cref="TokenValue"/> produces: plain text
/// in a member that could have held markup.
/// </summary>
/// <remarks>
/// A value read off a model never reaches a template as one. A value converter on both option sets
/// (<c>SharedFluid.BuildOptions</c>) swaps it for the string it wraps, so filters, conditions and
/// the substitution itself all see what a <c>string</c> member would have given them.
/// </remarks>
class TextToken(string? value) :
    TokenValue
{
    // Null when the string converted was: the operator is declared for a non-null string, but
    // nothing stops a null one arriving at run time.
    public string? Value { get; } = value;

    // For a token that reaches Fluid without passing the converter, wrapped in an ObjectValue by a
    // custom filter for instance. Fluid stringifies an object it has no other answer for, and the
    // default is the type name.
    public override string ToString() =>
        Value ?? string.Empty;
}
