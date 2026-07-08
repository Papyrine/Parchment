namespace Parchment;

/// <summary>
/// The editable-field values read from one document by <see cref="ParchmentExtractor"/>.
/// The document is a lossy projection of the model (loops expanded, formatting applied), so
/// extraction covers the <c>[EditableField]</c> subset only — <see cref="ApplyTo"/> merges that
/// subset onto a caller-supplied model instance.
/// </summary>
public sealed class ExtractResult<TModel>
{
    EditableMap map;

    internal ExtractResult(IReadOnlyList<ExtractedField> fields, EditableMap map)
    {
        Fields = fields;
        this.map = map;
        AllExtracted = fields.All(_ => _.State is FieldState.Extracted or FieldState.Empty);
    }

    /// <summary>
    /// One entry per editable member of <typeparamref name="TModel"/> (plus a
    /// <see cref="FieldState.Duplicate"/> entry per extra occurrence), in document order with
    /// missing members appended.
    /// </summary>
    public IReadOnlyList<ExtractedField> Fields { get; }

    /// <summary>
    /// True when every field read cleanly (<see cref="FieldState.Extracted"/> or
    /// <see cref="FieldState.Empty"/>) — no missing controls, parse failures, or duplicates.
    /// </summary>
    public bool AllExtracted { get; }

    /// <summary>
    /// Writes the extracted values onto <paramref name="model"/>:
    /// <see cref="FieldState.Extracted"/> values are assigned;
    /// <see cref="FieldState.Empty"/> assigns null to nullable members and leaves non-nullable
    /// members untouched; all other states are skipped. Reachability is validated first — if any
    /// applicable path has a null intermediate object, a <see cref="ParchmentExtractionException"/>
    /// is thrown before anything is mutated.
    /// </summary>
    public void ApplyTo(TModel model)
    {
        ArgumentNullException.ThrowIfNull(model);

        List<string>? unreachable = null;
        var applicable = new List<(EditableEntry Entry, object? Value)>();
        foreach (var field in Fields)
        {
            if (!map.TryGet(field.Path, out var entry))
            {
                continue;
            }

            var apply = field.State switch
            {
                FieldState.Extracted => true,
                FieldState.Empty when entry.IsNullable => true,
                _ => false
            };
            if (!apply)
            {
                continue;
            }

            if (entry.CanReach(model))
            {
                applicable.Add((entry, field.State == FieldState.Extracted ? field.Value : null));
            }
            else
            {
                (unreachable ??= []).Add(entry.DottedPath);
            }
        }

        if (unreachable != null)
        {
            throw new ParchmentExtractionException(
                $"Cannot apply extracted values to '{typeof(TModel).Name}': intermediate objects are null on path(s) {string.Join(", ", unreachable)}. Construct the target model with non-null intermediates first — nothing was applied.");
        }

        foreach (var (entry, value) in applicable)
        {
            entry.Setter(model, value);
        }
    }
}
