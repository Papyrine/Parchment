class OpenXmlMarkdownRenderer :
    RendererBase
{
    readonly Stack<ContainerState> stack = new();
    readonly Stack<ContainerState> pool = new();
    readonly List<(string TagName, Action<RunProperties> Apply)> activeInlineHtml = new();
    readonly Stack<int> indentStack = new();
    int currentIndent;
    int nextBookmarkId;
    LiteralInline? foldInline;
    Run? foldRun;
    Text? foldText;
    int unfoldedDepth;

    public OpenXmlMarkdownRenderer(MainDocumentPart mainPart, OpenXmlPart part, WordNumberingState numbering, ImagePolicies imagePolicies, BookmarkNames bookmarkNames, int headingOffset = 0)
    {
        MainPart = mainPart;
        Part = part;
        BookmarkNames = bookmarkNames;
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

        // Inline renderers. One for every inline MarkdigPipeline can produce: RendererBase writes
        // nothing for a leaf inline none of these claims, and raises nothing either, so a missing
        // one is text missing from the document. InlineCoverageTests holds the list to the pipeline.
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new HtmlEntityInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
        ObjectRenderers.Add(new DelimiterInlineRenderer());
        ObjectRenderers.Add(new AutolinkInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new HtmlInlineRenderer());
        ObjectRenderers.Add(new SmartyPantInlineRenderer());
    }

    public MainDocumentPart MainPart { get; }

    // The part the rendered elements land in, which is where a link's relationship has to be
    // added: a relationship id resolves against the part holding the element that carries it.
    // MainPart is still what OpenXmlHtml is handed, since that is all it takes - the relationships
    // it leaves there are moved by PartRelationships.
    public OpenXmlPart Part { get; }
    public int HeadingOffset { get; }

    public BookmarkNames BookmarkNames { get; }
    public WordNumberingState Numbering { get; }
    public ImagePolicies ImagePolicies { get; }

    // Tables flagged here had pipes aligned in the source across header, separator and body
    // rows — that pattern signals the user padded for readability, not custom widths, so the
    // table renderer skips emitting explicit column widths derived from dash counts. Populated
    // by MarkdownRendering before the renderer walks the AST.
    public HashSet<Markdig.Extensions.Tables.Table> SkipColumnWidths { get; } = [];

    // What the walk produced that a later whole-fragment pass exists for, so a render that produced
    // neither skips the pass instead of walking every element to find nothing. Each is set by the
    // only writers that can produce one: this renderer's own bookmark and section-marker writers,
    // and the two places converted html comes in, which is content this renderer did not build and
    // so cannot vouch for.
    public bool MayHoldBookmarks { get; set; }
    public bool MayHoldSections { get; set; }

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
        //
        // AddLiteral folds the one case that is safe before then, a literal directly after another,
        // so most of a bound value arrives here as one run already. Everything else - text beside an
        // entity, a soft line break or a smart quote - is still folded here.
        foreach (var run in top.CurrentRuns)
        {
            if (paragraph.LastChild is Run last &&
                SoleText(last) is { } previous &&
                SoleText(run) is { } incoming &&
                RunPropertiesComparer.Same(last.RunProperties, ((Run) run).RunProperties))
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
        MayHoldBookmarks = true;
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
        if (run is Run typed)
        {
            ApplyInlineHtmlFormats(typed);
        }

        stack.Peek().CurrentRuns.Add(run);
    }

    /// <summary>
    /// Adds the text of a literal, to the run of the literal before it when that is the inline
    /// directly before this one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Markdig ends a literal at every backslash escape and starts the next at the escaped
    /// character, so a bound value arrives as a chain of sibling literals with nothing between
    /// them. <see cref="FlushParagraph"/> folds the runs of such a chain into one, and this spares
    /// building them: two inlines that are adjacent siblings sit under the same emphasis, in the
    /// same link and between the same inline html tags, so whatever formats the run of the first
    /// is what would have formatted the run of the second.
    /// </para>
    /// <para>
    /// That is the whole of what makes it safe ahead of the flush. A literal is folded only into
    /// the run of its previous sibling, only while that run is still the last one written here,
    /// and only when both are a single <c>w:t</c> - a tab makes a run of several pieces. Anything
    /// else between two pieces of text is an inline of its own and ends the chain, and it is left
    /// to the flush as before.
    /// </para>
    /// </remarks>
    internal void AddLiteral(LiteralInline inline, CharSpan text)
    {
        var hasTab = text.Contains('\t');
        if (foldText != null &&
            !hasTab &&
            ReferenceEquals(foldInline, inline.PreviousSibling) &&
            stack.Peek().CurrentRuns is [.., var last] &&
            ReferenceEquals(last, foldRun))
        {
            foldText.Text = string.Concat(foldText.Text, text);
            foldInline = inline;
            return;
        }

        var run = LiteralInlineRenderer.BuildRun(text.ToString());
        AddRun(run);
        if (hasTab ||
            unfoldedDepth > 0)
        {
            foldInline = null;
            foldRun = null;
            foldText = null;
            return;
        }

        // The text is the last child: AddRun puts an rPr ahead of it when an inline html tag is open.
        foldInline = inline;
        foldRun = run;
        foldText = (Text) run.LastChild!;
    }

    /// <summary>
    /// Keeps the literals written until <see cref="EndUnfoldedRuns"/> a run each.
    /// </summary>
    /// <remarks>
    /// For runs that leave <see cref="ContainerState.CurrentRuns"/> before the paragraph is flushed:
    /// the text of a link is moved into its <c>w:hyperlink</c>, where <see cref="FlushParagraph"/>
    /// does not look, so those runs were never folded and folding them here would change the output.
    /// </remarks>
    internal void BeginUnfoldedRuns() =>
        unfoldedDepth++;

    internal void EndUnfoldedRuns() =>
        unfoldedDepth--;

    /// <summary>
    /// Gives <paramref name="run"/> the formatting of every inline html tag open around it.
    /// </summary>
    /// <remarks>
    /// <see cref="AddRun"/> does this for a run added directly. A renderer that adds its runs
    /// already wrapped - an autolink arrives as one <c>Hyperlink</c> - calls it for the runs inside,
    /// which <see cref="AddRun"/> does not look for.
    /// </remarks>
    internal void ApplyInlineHtmlFormats(Run run)
    {
        if (activeInlineHtml.Count == 0)
        {
            return;
        }

        var properties = run.RunProperties ??= new();
        foreach (var (_, apply) in activeInlineHtml)
        {
            apply(properties);
        }
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

    /// <summary>
    /// Drop every format still open. An unclosed tag would otherwise apply to every run in the
    /// rest of the document, since nothing else ever removes it.
    /// </summary>
    internal void ClearInlineHtmlFormats() =>
        activeInlineHtml.Clear();

    internal void AddBlock(OpenXmlElement block)
    {
        FlushParagraph();
        var blocks = stack.Peek().Blocks;

        // Word merges two w:tbl siblings into one table, so a table that would directly follow
        // another gets an empty paragraph first. Every table passes through here - a pipe or grid
        // table, and each one an html block converts to - and a list item or block quote writes
        // into the container around it, so this one check covers the body, both of those and a
        // grid-table cell, which is a container of its own.
        if (block is Table &&
            AdjacentTables.EndsWithTable(blocks))
        {
            blocks.Add(new Paragraph());
        }

        blocks.Add(block);
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
