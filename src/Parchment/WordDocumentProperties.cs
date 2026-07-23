namespace Parchment;

/// <summary>
/// Standard and user-defined document properties — the values Word surfaces in the File &gt; Info
/// pane and the Advanced Properties dialog. Pass to <see cref="TemplateStore.Render(string, object, Stream, WordDocumentProperties, Cancel)"/>.
/// </summary>
/// <remarks>
/// Every member is optional and only the values that are set are written, so a property left at its
/// default leaves that part of the document as the template had it. This matters because a template
/// usually arrives carrying properties of its own: the parts are merged, never rewritten wholesale.
/// Mirrors <c>Excelsior.WordDocumentProperties</c>.
/// </remarks>
public class WordDocumentProperties
{
    /// <summary>Maps to the core property <c>Title</c>.</summary>
    public string? Title { get; init; }

    /// <summary>The document author. Maps to the core property <c>Creator</c> (<c>dc:creator</c>).</summary>
    public string? Author { get; init; }

    /// <summary>Maps to the core property <c>Subject</c>.</summary>
    public string? Subject { get; init; }

    /// <summary>Comma- or space-separated tags. Maps to the core property <c>Keywords</c>.</summary>
    public string? Keywords { get; init; }

    /// <summary>Free-text comments. Maps to the core property <c>Description</c> (<c>dc:description</c>).</summary>
    public string? Comments { get; init; }

    /// <summary>Maps to the core property <c>Category</c>.</summary>
    public string? Category { get; init; }

    /// <summary>Content/approval status (e.g. "Draft", "Final"). Maps to the core property <c>ContentStatus</c>.</summary>
    public string? Status { get; init; }

    /// <summary>Maps to the core property <c>LastModifiedBy</c>.</summary>
    public string? LastModifiedBy { get; init; }

    /// <summary>Company name. Written to the extended (app) properties part.</summary>
    public string? Company { get; init; }

    /// <summary>Manager name. Written to the extended (app) properties part.</summary>
    public string? Manager { get; init; }

    /// <summary>
    /// User-defined properties shown on the "Custom" tab of Word's Advanced Properties dialog.
    /// Supported value types are <see cref="string"/>, <see cref="bool"/>, integral and
    /// floating-point numbers, <see cref="DateTime"/>, <see cref="DateOnly"/> and <see cref="Guid"/>;
    /// any other value type throws an <see cref="ArgumentException"/> at render, so convert it to a
    /// supported type first. A <c>null</c> value is written as empty text.
    /// </summary>
    /// <remarks>
    /// Entries here are merged over whatever the template already defines: a name that matches an
    /// existing property replaces it, and one that does not is added. Properties the template
    /// carries and this dictionary does not name are left alone. To remove one, use
    /// <see cref="RemoveCustom"/>.
    /// </remarks>
    public Dictionary<string, object?> Custom { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Names of user-defined properties to drop from the template's own set. Applied after
    /// <see cref="Custom"/>, so a name in both is removed.
    /// </summary>
    public HashSet<string> RemoveCustom { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Discards the properties the template carries before the values above are applied, so the
    /// output starts empty rather than inheriting them.
    /// </summary>
    /// <remarks>
    /// A template that a person has edited carries their editing history: <c>Creator</c> and
    /// <c>LastModifiedBy</c> name them, and <c>Revision</c> and <c>LastPrinted</c> describe work
    /// that has nothing to do with the generated document. Merging stays the default because a
    /// template's properties are usually wanted — set this where they are not, and a document
    /// rendered from someone's copy of the template stops carrying their name.
    ///
    /// Covers the core part in full, and <see cref="Company"/> and <see cref="Manager"/> on the
    /// extended part. User-defined properties are left alone, since those are normally the
    /// template's own data rather than metadata about who edited it — use
    /// <see cref="RemoveCustom"/> to drop those by name.
    /// </remarks>
    public bool ClearBuiltIn { get; init; }
}
