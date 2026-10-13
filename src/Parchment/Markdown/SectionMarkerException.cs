/// <summary>
/// Thrown from the markdown walk when a <c>[SECTION]</c> marker names a page setup Word cannot hold.
/// The walk cannot see the template name, so this is caught in
/// <c>RegisteredMarkdownTemplate.Render</c> and rethrown as a <see cref="ParchmentRenderException"/>.
/// In the docx flow the token that carried the markdown wraps it, as it wraps any other failure.
/// </summary>
sealed class SectionMarkerException(string message) :
    Exception(message);
