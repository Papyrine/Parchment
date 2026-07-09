namespace Parchment;

/// <summary>
/// One editable collection reconstructed from a repeating section: <see cref="List"/> is the rebuilt
/// list (a <c>List&lt;TItem&gt;</c> or <c>TItem[]</c> matching the model member) when
/// <see cref="State"/> is <see cref="FieldState.Extracted"/>, or null when the repeating-section
/// container was not found (<see cref="FieldState.Missing"/>). <see cref="ExtractResult{TModel}.ApplyTo"/>
/// assigns <see cref="List"/> onto the collection member — replace-all.
/// </summary>
sealed record ExtractedCollection(
    CollectionEntry Entry,
    FieldState State,
    object? List);
