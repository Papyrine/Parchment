namespace Parchment.Generated;

public sealed record FormatMapEntry(string DottedPath, FormatMapKind Format, Func<object, object?> Getter);