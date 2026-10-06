/// <summary>
/// Renders a value that produces Word content rather than markdown source — an <c>| html</c>
/// value, an <see cref="HtmlToken"/> or an <see cref="OpenXmlToken"/> — in the markdown flow.
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
/// filter writes a marker and parks the value on the render's ambient values, and once Markdig has
/// finished the marker's paragraph is swapped for what the value produces. Registration cannot do
/// this the way <c>MarkdownFormats</c> does — the value only exists at render — so the pairing is
/// carried on the <see cref="TemplateContext"/> that the render owns.
/// </para>
/// <para>
/// An <see cref="OpenXmlToken"/> takes the same route for the same reason. It has no markdown form
/// at all, but by the time the marker is swapped there is a document to emit into and a host
/// paragraph to emit in place of — which is everything <see cref="IOpenXmlContext"/> needs. This is
/// the flow already emitting raw OpenXML for <c>[ExcelsiorTable]</c> members; the token is the same
/// mechanism opened to a caller.
/// </para>
/// </remarks>
static class MarkdownTokenBlocks
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
        public Dictionary<string, TokenValue> Pending { get; } = new(StringComparer.Ordinal);

        public void Dispose() =>
            current.Value = null;
    }

    static readonly AsyncLocal<Scope?> current = new();

    /// <summary>
    /// Opens the window in which <see cref="Register(TokenValue)"/> can run. Disposing closes it; the returned
    /// scope keeps the map, which outlives the window.
    /// </summary>
    public static Scope BeginScope()
    {
        var scope = new Scope();
        current.Value = scope;
        return scope;
    }

    /// <summary>
    /// Parks html source, as the <c>| html</c> filter has only a string to hand.
    /// </summary>
    public static string Register(string html) =>
        Register(new HtmlToken(html));

    /// <summary>
    /// Parks <paramref name="token"/> and returns the marker standing in for it. One per call rather
    /// than per member: a token inside a loop renders once per iteration, each with its own value.
    /// </summary>
    public static string Register(TokenValue token)
    {
        var scope = current.Value ??
                    throw new("A token was registered outside a markdown render. Only the markdown flow parks a value for later conversion, and only RegisteredMarkdownTemplate opens the scope it is parked in.");

        // Wrapped in the private-use sentinel MarkdownExcelsiorTables uses, for the two reasons it
        // uses it. No rendered value can contain one, so a value that happened to read like a
        // marker is not swapped for someone else's content. And it terminates the index: Apply reads
        // a marker as what lies between two sentinels, so an unterminated "…-1" could not be told
        // from the start of "…-10".
        var marker = $"{sentinel}parchment-token-{scope.Pending.Count}{sentinel}";
        scope.Pending[marker] = token;
        return marker;
    }

    /// <summary>
    /// Swaps each marker's paragraph for what the value it stands for produces.
    /// </summary>
    public static void Apply(
        Body body,
        IReadOnlyDictionary<string, TokenValue> pending,
        MainDocumentPart mainPart,
        WordNumberingState numbering,
        Lazy<StyleSet> styles,
        ImagePolicies imagePolicies,
        string templateName)
    {
        if (pending.Count == 0)
        {
            return;
        }

        // Every host is found in one walk of the body, before anything is swapped. Looking for each
        // marker with its own walk made the cost the number of markers times the number of
        // paragraphs, and a token inside a loop adds to both with every iteration - so a report
        // twice as long took four times as long. Finding them all first is safe because nothing a
        // swap inserts can hold a marker: a marker only ever comes out of the liquid render, which
        // is over. Materialized before mutating for the reason it always was - the swap edits the
        // tree being walked.
        var hosts = new Dictionary<string, List<Paragraph>>(StringComparer.Ordinal);
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            foreach (var marker in Markers(paragraph.InnerText, pending))
            {
                if (!hosts.TryGetValue(marker, out var list))
                {
                    hosts[marker] = list = [];
                }

                list.Add(paragraph);
            }
        }

        // In the order the markers were registered, so a block holding two of them is still
        // reported against the first. A marker inside a loop iteration that rendered nothing, or
        // inside a false conditional, has no host at all, which is not an error.
        var replaced = new List<Paragraph>();
        foreach (var (marker, token) in pending)
        {
            if (!hosts.TryGetValue(marker, out var list))
            {
                continue;
            }

            foreach (var host in list)
            {
                Insert(host, marker, token, mainPart, numbering, styles, imagePolicies, templateName);
                replaced.Add(host);
            }
        }

        Remove(replaced);
    }

    /// <summary>
    /// Takes the hosts out once everything they stood for is in, a parent at a time.
    /// </summary>
    /// <remarks>
    /// OpenXML links an element's children forwards only, so removing one walks every sibling before
    /// it to find the one to relink. Removing each host as it was swapped therefore cost the length
    /// of the body per host - the cost the single walk in <see cref="Apply"/> is there to avoid,
    /// cheaper per step but growing the same way. A parent's children are instead detached and
    /// re-attached once, without the hosts, which is constant per child: detaching always takes the
    /// first child and attaching always follows the last.
    /// </remarks>
    static void Remove(List<Paragraph> hosts)
    {
        var gone = new HashSet<OpenXmlElement>(hosts);

        // Materialized before any parent is emptied: a host's Parent is null from then on.
        var parents = hosts
            .Select(_ => _.Parent!)
            .Distinct()
            .ToList();

        foreach (var parent in parents)
        {
            var kept = new List<OpenXmlElement>();
            for (var child = parent.FirstChild; child != null; child = child.NextSibling())
            {
                if (!gone.Contains(child))
                {
                    kept.Add(child);
                }
            }

            parent.RemoveAllChildren();
            foreach (var child in kept)
            {
                parent.AppendChild(child);
            }
        }
    }

    // The markers of this render that text holds, each once. Almost every paragraph holds none, and
    // is dismissed on the search for a sentinel.
    static List<string> Markers(string text, IReadOnlyDictionary<string, TokenValue> pending)
    {
        var markers = new List<string>();
        var start = text.IndexOf(sentinel);
        while (start >= 0)
        {
            var end = text.IndexOf(sentinel, start + 1);
            if (end < 0)
            {
                break;
            }

            var candidate = text.Substring(start, end - start + 1);
            if (!pending.ContainsKey(candidate))
            {
                // Not one of these markers - an [ExcelsiorTable] placeholder is wrapped the same
                // way - so its closing sentinel may be the opening one of a marker that is.
                start = end;
                continue;
            }

            if (!markers.Contains(candidate))
            {
                markers.Add(candidate);
            }

            start = text.IndexOf(sentinel, end + 1);
        }

        return markers;
    }

    // Private-use, so no rendered value can contain one - see Register. Written as its code point
    // because the character has no glyph, and a literal holding it reads as an empty one.
    const char sentinel = (char) 0xE000;

    // Puts what the value produces after its host and leaves the host where it is - see Remove.
    // After rather than before, because inserting after an element is constant where inserting
    // before one walks every sibling ahead of it, the same way removing one does.
    static void Insert(
        Paragraph host,
        string marker,
        TokenValue token,
        MainDocumentPart mainPart,
        WordNumberingState numbering,
        Lazy<StyleSet> styles,
        ImagePolicies imagePolicies,
        string templateName)
    {
        if (host.InnerText.Trim() != marker)
        {
            // These convert to whole blocks — paragraphs, lists, tables — so the value replaces the
            // block it sits in, and text sharing that block would be discarded. Refused rather than
            // silently dropped, matching how an [ExcelsiorTable] token answers the same mistake.
            throw new ParchmentRenderException(
                templateName,
                $"A '{Describe(token)}' converts to Word blocks that replace the block it sits in, " +
                "so it has to be alone in that block — the text sharing it would be discarded. " +
                $"Markdown parsed it into a paragraph reading: {host.InnerText.Trim()}");
        }

        var parent = host.Parent!;
        OpenXmlElement cursor = host;
        foreach (var element in Render(token, host, mainPart, numbering, styles, imagePolicies))
        {
            cursor = parent.InsertAfter(element, cursor);
        }
    }

    static IEnumerable<OpenXmlElement> Render(
        TokenValue token,
        Paragraph host,
        MainDocumentPart mainPart,
        WordNumberingState numbering,
        Lazy<StyleSet> styles,
        ImagePolicies imagePolicies) =>
        token switch
        {
            HtmlToken html => WordHtmlConverter.ToElements(html.Source, mainPart, imagePolicies.BuildSettings()),
            OpenXmlToken raw when ReferenceEquals(raw, OpenXmlToken.Empty) => [],
            // The host is handed over as well as the part: a token that only wants to know what it
            // is replacing - its style, its indentation - gets the same answer it would in the docx
            // flow, even though the paragraph here was built by Markdig rather than by a template.
            OpenXmlToken raw => raw.Render(new OpenXmlContextImpl(mainPart, numbering, styles.Value, host)),
            _ => []
        };

    static string Describe(TokenValue token)
    {
        if (token is HtmlToken)
        {
            return "| html";
        }

        return nameof(OpenXmlToken);
    }
}
