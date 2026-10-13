/// <summary>
/// The hint name a target gets when its flattened name is shared with another target. See
/// <see cref="HintNames.Disambiguate"/>.
/// </summary>
sealed record HintOverride(string ModelFullyQualifiedName, string HintName);
