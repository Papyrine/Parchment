class OpenXmlMarkdownRenderer :
    RendererBase
{
    readonly Stack<ContainerState> stack = new();
    readonly Stack<ContainerState> pool = new();
    readonly List<(string TagName, Action<RunProperties> Apply)> activeInlineHtml = new();
    readonly Stack<int> indentStack = new();
    int currentIndent;
    int nextBookmarkId;

    public OpenXmlMarkdownRenderer(MainDocumentPart mainPart, WordNumberingState numbering, ImagePolicies imagePolicies, int headingOffset = 0)
    {
        MainPart = mainPart;
        HeadingOffset = headingOffset;
        Numbering = numbering;
        ImagePolicies = imagePolicies;
        stack.Push(new());

        // Block renderers
        ObjectRenderers.Add(new HeadingBlockRenderer());
        ObjectRenderers.Add(new ParagraphBlockRenderer());
        ObjectRenderers.Add(new ListBlockRenderer());
        ObjectRenderers.Add(new QuoteBlockRenderer());
        ObjectRenderers.Add(new TableRenderer());
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ThematicBreakRenderer());
        ObjectRenderers.Add(new HtmlBlockRenderer());

        // Inline renderers
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
        ObjectRenderers.Add(new AutolinkInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new HtmlInlineRenderer());
        ObjectRenderers.Add(new SmartyPantInlineRenderer());
    }

    public MainDocumentPart MainPart { get; }
    public int HeadingOffset { get; }
    public WordNumberingState Numbering { get; }
    public ImagePolicies ImagePolicies { get; }

    // Tables flagged here had pipes aligned in the source across header, separator and body
    // rows — that pattern signals the user padded for readability, not custom widths, so the
    // table renderer skips emitting explicit column widths derived from dash counts. Populated
    // by MarkdownRendering before the renderer walks the AST.
    public HashSet<Markdig.Extensions.Tables.Table> SkipColumnWidths { get; } = [];

    internal ContainerState Top => stack.Peek();

    public override object Render(MarkdownObject markdownObject)
    {
        Write(markdownObject);
        return this;
    }

    public IReadOnlyList<OpenXmlElement> Drain() =>
        stack.Peek().Blocks;

    internal void PushContainer() =>
        stack.Push(pool.Count > 0 ? pool.Pop() : new());

    internal ContainerState PopContainer() =>
        stack.Pop();

    /// <summary>
    /// Return a popped <see cref="ContainerState"/> to the pool for reuse. Lists are cleared
    /// so subsequent <see cref="PushContainer"/> calls receive an empty state. For high-cell
    /// markdown tables this avoids one ContainerState + two List allocations per cell.
    /// </summary>
    internal void ReleaseContainer(ContainerState state)
    {
        state.Blocks.Clear();
        state.CurrentRuns.Clear();
        pool.Push(state);
    }

    internal void FlushParagraph(ParagraphProperties? properties = null)
    {
        var top = stack.Peek();
        if (top.CurrentRuns.Count == 0 && properties == null)
        {
            return;
        }

        var paragraph = new Paragraph();
        if (properties != null)
        {
            paragraph.ParagraphProperties = properties;
        }

        // Markdig hands adjacent text out as separate inlines whenever anything interrupts the
        // literal — a backslash escape most of all, and MarkdownEncoder emits one for every
        // punctuation character in a bound value. A run each would triple the run count of ordinary
        // prose ("on-call" alone becoming two) for no visible difference, so identically-formatted
        // neighbours are folded back into one.
        //
        // Here rather than in AddRun: an inline renderer can still reach back and mutate the runs it
        // produced after adding them — EmphasisInlineRenderer applies its formatting to everything
        // written since it started — so a merge at add time would fuse a run into its neighbour
        // before that formatting landed, and would leave the renderer's index range pointing at the
        // wrong runs. By the time a paragraph flushes, every inline in it is final.
        foreach (var run in top.CurrentRuns)
        {
            if (paragraph.LastChild is Run last &&
                SoleText(last) is { } previous &&
                SoleText(run) is { } incoming &&
                SameProperties(last, (Run) run))
            {
                previous.Text += incoming.Text;
                continue;
            }

            paragraph.Append(run);
        }

        top.Blocks.Add(paragraph);
        top.CurrentRuns.Clear();
    }

    /// <summary>
    /// Open a bookmark around whatever is written next in the current paragraph.
    /// </summary>
    /// <remarks>
    /// Ids are allocated per render and renumbered document-wide once the walk finishes (see
    /// <see cref="MarkdownRendering"/>), because html blocks bring their own bookmarks with their own
    /// numbering and the two sequences would otherwise overlap.
    /// </remarks>
    internal string AddBookmarkStart(string name)
    {
        var id = (++nextBookmarkId).ToString(CultureInfo.InvariantCulture);
        AddRun(
            new BookmarkStart
            {
                Id = id,
                Name = name
            });
        return id;
    }

    internal void AddBookmarkEnd(string id) =>
        AddRun(
            new BookmarkEnd
            {
                Id = id
            });

    internal void AddRun(OpenXmlElement run)
    {
        if (run is Run typed && activeInlineHtml.Count > 0)
        {
            typed.RunProperties ??= new();
            foreach (var (_, apply) in activeInlineHtml)
            {
                apply(typed.RunProperties);
            }
        }

        stack.Peek().CurrentRuns.Add(run);
    }

    /// <summary>
    /// The single <c>w:t</c> a run is made of, or null when it is anything else — a bookmark, a
    /// break, an image, or a run carrying tabs. Only the single-text case can be merged: everything
    /// else has structure between or around the text that concatenation would lose.
    /// </summary>
    static Text? SoleText(OpenXmlElement element)
    {
        if (element is not Run run)
        {
            return null;
        }

        Text? text = null;
        foreach (var child in run.ChildElements)
        {
            if (child is RunProperties)
            {
                continue;
            }

            if (text != null ||
                child is not Text candidate)
            {
                return null;
            }

            text = candidate;
        }

        return text;
    }

    static bool SameProperties(Run left, Run right)
    {
        var leftProperties = left.RunProperties;
        var rightProperties = right.RunProperties;

        if (leftProperties == null)
        {
            return rightProperties == null;
        }

        return rightProperties != null &&
               leftProperties.OuterXml == rightProperties.OuterXml;
    }

    internal void PushInlineHtmlFormat(string tagName, Action<RunProperties> apply) =>
        activeInlineHtml.Add((tagName, apply));

    internal void PopInlineHtmlFormat(string tagName)
    {
        for (var i = activeInlineHtml.Count - 1; i >= 0; i--)
        {
            if (string.Equals(activeInlineHtml[i].TagName, tagName, StringComparison.OrdinalIgnoreCase))
            {
                activeInlineHtml.RemoveAt(i);
                return;
            }
        }
    }

    internal void AddBlock(OpenXmlElement block)
    {
        FlushParagraph();
        stack.Peek().Blocks.Add(block);
    }

    internal int CurrentIndent => currentIndent;

    internal void PushIndent(int dxa)
    {
        indentStack.Push(dxa);
        currentIndent += dxa;
    }

    internal void PopIndent() =>
        currentIndent -= indentStack.Pop();
}