sealed record TypeEntry(
    string TypeFullyQualifiedName,
    string? ElementTypeFullyQualifiedName,
    EquatableArray<MemberEntry> Members,
    // Whether the type has a public parameterless constructor — consulted by the editable
    // collection shape rules (extraction rebuilds the list, so elements must be constructable).
    bool HasParameterlessCtor = false,
    // What Fluid hands a template for a value of this type. Anything but None answers to liquid's
    // own members instead of the ones above — see ShapeResolver.TryResolveLiquidMember.
    LiquidKind Liquid = LiquidKind.None);
