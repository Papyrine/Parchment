/// <summary>
/// Turns whatever a render threw into the <see cref="ParchmentRenderException"/> the caller was
/// promised, with the context known at the point that caught it.
/// </summary>
/// <remarks>
/// Used as an exception filter plus a wrap, at several depths: the innermost catch has the token or
/// tag, the outermost only the template name, and <see cref="ShouldWrap"/> lets an exception that was
/// already wrapped further in pass every catch above it untouched.
/// </remarks>
static class RenderFailure
{
    /// <summary>
    /// False for an exception that already carries its context, and for a cancellation, which the
    /// caller has to be able to recognise as one.
    /// </summary>
    public static bool ShouldWrap(Exception exception) =>
        Cause(exception) is not (ParchmentException or OperationCanceledException);

    public static ParchmentRenderException Wrap(
        Exception exception,
        string templateName,
        string? partUri = null,
        string? contextSnippet = null,
        string? tokenSource = null)
    {
        var cause = Cause(exception);
        return new(
            templateName,
            cause.Message,
            partUri,
            contextSnippet,
            tokenSource,
            inner: cause);
    }

    // A reflection call reports what its target threw as a TargetInvocationException, which says
    // nothing about the failure.
    static Exception Cause(Exception exception)
    {
        while (exception is TargetInvocationException { InnerException: { } inner })
        {
            exception = inner;
        }

        return exception;
    }
}
