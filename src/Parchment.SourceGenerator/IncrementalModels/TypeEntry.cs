sealed record TypeEntry(
    string TypeFullyQualifiedName,
    string? ElementTypeFullyQualifiedName,
    EquatableArray<MemberEntry> Members,
    // Whether the type has a public parameterless constructor — consulted by the editable
    // collection shape rules (extraction rebuilds the list, so elements must be constructable).
    bool HasParameterlessCtor = false);
