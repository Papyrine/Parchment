namespace Parchment;

/// <summary>
/// Raised by <see cref="ParchmentExtractor"/> when a document cannot be read (not a docx, no
/// main body), when the model declares no editable members, or by
/// <c>ExtractResult&lt;TModel&gt;.ApplyTo</c> when extracted values cannot be written back
/// because intermediate objects on a dotted path are null.
/// </summary>
public sealed class ParchmentExtractionException(string message, Exception? inner = null) :
    ParchmentException(message, inner);
