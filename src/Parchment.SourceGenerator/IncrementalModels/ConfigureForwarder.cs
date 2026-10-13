/// <summary>
/// A method the generator adds to a partial type so the model's registration can reach an
/// <c>[ExcelsiorTable(Configure = ...)]</c> method it could not call directly: a private one on a
/// type other than the model. <see cref="Types"/> is the declaring type and everything enclosing
/// it, outermost first, each name carrying its type parameters.
/// </summary>
sealed record ConfigureForwarder(
    string? Namespace,
    EquatableArray<EnclosingType> Types,
    string Name,
    string ParameterType,
    string Method);
