class OpenXmlMarkdownRenderer :
    RendererBase
{
    readonly Stack<ContainerState> stack = new();
    readonly Stack<ContainerState> pool = new();
    readonly List<(string TagName, Action<RunProperties> Apply)> activeInlineHtml = new();
    readonly Stack<int> indentStack = new();
    int currentIndent;

    public OpenXmlMarkdownRenderer(MainDocumentPart mainPart, WordNumberingState numbering, ImagePolicies imagePolicies, int headingOffset = 0)
    {
        MainPart = mainPart;
        HeadingOffset = headingOffset;
        AvailableStyles = StyleSet.Read(mainPart);
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
    public StyleSet AvailableStyles { get; }
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

        foreach (var run in top.CurrentRuns)
        {
            paragraph.Append(run);
        }

        top.Blocks.Add(paragraph);
        top.CurrentRuns.Clear();
    }

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