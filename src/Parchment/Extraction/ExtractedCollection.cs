namespace Parchment;

/// <summary>
/// One editable collection read from a repeating section: <see cref="Rows"/> holds the extracted
/// values of each kept item, in document order, when <see cref="State"/> is
/// <see cref="FieldState.Extracted"/>, or null when the repeating-section container was not found
/// (<see cref="FieldState.Missing"/>). <see cref="ExtractResult{TModel}.ApplyTo"/> assigns
/// <see cref="Build"/> onto the collection member — replace-all.
/// </summary>
sealed record ExtractedCollection(
    CollectionEntry Entry,
    FieldState State,
    List<List<(EditableEntry Entry, object? Value)>>? Rows)
{
    /// <summary>
    /// Builds the collection (a <c>List&lt;TItem&gt;</c> or <c>TItem[]</c> matching the model
    /// member) with newly constructed items, so every model it is applied to owns its own.
    /// </summary>
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public object Build()
    {
        var listType = typeof(List<>).MakeGenericType(Entry.ElementType);
        var list = (IList)Activator.CreateInstance(listType)!;

        foreach (var row in Rows!)
        {
            var item = Entry.ElementFactory();
            foreach (var (entry, value) in row)
            {
                entry.Setter(item, value);
            }

            list.Add(item);
        }

        if (!Entry.IsArray)
        {
            return list;
        }

        var array = Array.CreateInstance(Entry.ElementType, list.Count);
        list.CopyTo(array, 0);
        return array;
    }
}
