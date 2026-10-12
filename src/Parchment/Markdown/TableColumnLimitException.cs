/// <summary>
/// Thrown by <see cref="TableRenderer"/> for a markdown table wider than Word can hold. The
/// renderer cannot see the template name, nor in the docx flow the token the markdown arrived
/// through, so each caller of <see cref="MarkdownRendering"/> catches this and rethrows it as a
/// <see cref="ParchmentRenderException"/>.
/// </summary>
sealed class TableColumnLimitException(int columns, int limit) :
    Exception(
        $"A markdown table is {columns} columns wide at its widest row, and a Word table holds at most {limit}. " +
        "Word does not open a document that holds a wider one, so split the table or remove columns.");
