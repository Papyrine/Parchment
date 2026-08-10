/// <summary>
/// Converts an <c>| html</c> value in the markdown flow with the html converter, rather than
/// leaving markdown to guess what it is.
/// </summary>
/// <remarks>
/// <para>
/// The docx flow turns the value into an <c>HtmlToken</c> and hands it to
/// <c>OpenXmlHtml.WordHtmlConverter</c>, so what comes out is decided by html rules. A markdown
/// template renders liquid to text first, so a value written straight into that text is not html
/// yet — it is source that Markdig will classify. That classification is not the same thing:
/// <c>&lt;span&gt;a *b* c&lt;/span&gt;</c> written into markdown emits the span as literal text and
/// turns <c>*b*</c> italic, which is markdown's answer, not html's.
/// </para>
/// <para>
/// So the value takes the same route the tables do (see <see cref="MarkdownExcelsiorTables"/>): the
/// filter writes a marker and parks the source on the render's ambient values, and once Markdig has
/// finished the marker's paragraph is swapped for the converter's output. Registration cannot do
/// this the way <c>MarkdownFormats</c> does — the value only exists at render — so the pairing is
/// carried on the <see cref="TemplateContext"/> that the render owns.
/// </para>
/// </remarks>
static class MarkdownHtmlBlocks
{
    /// <summary>
    /// The markers registered by one markdown render, and what each stands for.
    /// </summary>
    /// <remarks>
    /// Held in an <see cref="AsyncLocal{T}"/> because the two callers cannot both be handed one.
    /// The <c>html</c> filter is given the <see cref="TemplateContext"/> and could park the source
    /// on its ambient values; the <see cref="HtmlToken"/> path cannot, because it runs from a Fluid
    /// value converter, whose signature is <c>Func&lt;object, object&gt;</c> — no context, and
    /// <c>FluidValue.WriteTo</c> is handed only an encoder and a culture. One mechanism both can
    /// reach beats two that disagree, which is what left <c>HtmlToken</c> passing through while the
    /// filter converted.
    ///
    /// The scope covers the liquid render and nothing more: <c>RegisteredMarkdownTemplate</c> takes
    /// the map off it as soon as Fluid is done and hands it to <see cref="Apply"/> explicitly, so
    /// the ambient window is the one place it is unavoidable. AsyncLocal rather than ThreadStatic
    /// because the render is async — the flow, not the thread, is what a render owns.
    /// </remarks>
    public sealed class Scope :
        IDisposable
    {
        public Dictionary<string, string> Pending { get; } = new(StringComparer.Ordinal);

        public void Dispose() =>
            current.Value = null;
    }

    static readonly AsyncLocal<Scope?> current = new();

    /// <summary>
    /// Opens the window in which <see cref="Register"/> can run. Disposing closes it; the returned
    /// scope keeps the map, which outlives the window.
    /// </summary>
    public static Scope BeginScope()
    {
        var scope = new Scope();
        current.Value = scope;
        return scope;
    }

    /// <summary>
    /// Parks <paramref name="html"/> and returns the marker standing in for it. One per call rather
    /// than per member: a token inside a loop renders once per iteration, each with its own value.
    /// </summary>
    public static string Register(string html)
    {
        var scope = current.Value ??
                    throw new("Html was registered outside a markdown render. Only the markdown flow parks html for later conversion, and only RegisteredMarkdownTemplate opens the scope it is parked in.");

        // Wrapped in the private-use sentinel MarkdownExcelsiorTables uses, for the two reasons it
        // uses it. No rendered value can contain one, so a value that happened to read like a
        // marker is not swapped for someone else's html. And it terminates the index: Apply locates
        // a host with Contains, so an unterminated "…-1" would be found inside "…-10" and the
        // first marker reported as sharing its block with what is really the eleventh's.
        var marker = $"parchment-html-{scope.Pending.Count}";
        scope.Pending[marker] = html;
        return marker;
    }

    /// <summary>
    /// Swaps each marker's paragraph for the converted html.
    /// </summary>
    public static void Apply(
        Body body,
        IReadOnlyDictionary<string, string> pending,
        MainDocumentPart mainPart,
        ImagePolicies imagePolicies,
        string templateName)
    {
        if (pending.Count == 0)
        {
            return;
        }

        foreach (var (marker, html) in pending)
        {
            // Materialized before mutating: the swap edits the tree being walked. A marker inside a
            // loop iteration that rendered nothing, or inside a false conditional, has no host at
            // all, which is not an error.
            var hosts = body.Descendants<Paragraph>()
                .Where(_ => _.InnerText.Contains(marker, StringComparison.Ordinal))
                .ToList();

            foreach (var host in hosts)
            {
                Replace(host, marker, html, mainPart, imagePolicies, templateName);
            }
        }
    }

    static void Replace(
        Paragraph host,
        string marker,
        string html,
        MainDocumentPart mainPart,
        ImagePolicies imagePolicies,
        string templateName)
    {
        if (host.InnerText.Trim() != marker)
        {
            // Html converts to whole blocks — paragraphs, lists, tables — so it replaces the block
            // it sits in, and text sharing that block would be discarded. Refused rather than
            // silently dropped, matching how an [ExcelsiorTable] token answers the same mistake.
            throw new ParchmentRenderException(
                templateName,
                "An '| html' token converts to Word blocks that replace the block it sits in, so it " +
                "has to be alone in that block — the text sharing it would be discarded. Markdown " +
                $"parsed it into a paragraph reading: {host.InnerText.Trim()}");
        }

        var parent = host.Parent!;
        var elements = WordHtmlConverter.ToElements(html, mainPart, imagePolicies.BuildSettings());
        foreach (var element in elements)
        {
            parent.InsertBefore(element, host);
        }

        host.Remove();
    }
}
