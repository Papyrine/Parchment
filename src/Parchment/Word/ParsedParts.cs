/// <summary>
/// The token-bearing parts of a registered docx template, parsed once and shared by every render.
/// </summary>
/// <remarks>
/// Parsing those parts from xml was a third of a render. They are the same for every render, so
/// they are parsed once, the first time the template is rendered, and each render deep-clones the
/// elements into a package whose copies of those parts hold an empty root - the docx counterpart of
/// what <c>TemplateStore.BakeStyleSource</c> does for the markdown flow.
///
/// The elements are read back from the registration snapshot and never taken from the document
/// registration scanned: a render has always worked on a parse of the saved bytes, and that is not
/// the tree that was saved (a retyped <c>w:smartTag</c> does not survive a save, and neither does
/// an annotation). Cloning a parse of the same bytes gives a render the tree it has always had,
/// which is what keeps the output the same bytes.
///
/// The kept elements are read by concurrent renders with no lock, so nothing may change them once
/// <see cref="Read"/> returns: no render is ever handed one, only a clone. That includes the
/// loading the Open XML SDK defers - an element holding unparsed xml parses it the first time its
/// children are asked for, which a clone does - so every element is walked here, before the tree
/// is shared.
/// </remarks>
sealed class ParsedParts
{
    readonly Dictionary<string, OpenXmlElement[]> children;

    ParsedParts(byte[] bytes, Dictionary<string, OpenXmlElement[]> children)
    {
        Bytes = bytes;
        this.children = children;
    }

    /// <summary>
    /// The registration snapshot with the root of every token-bearing part emptied.
    /// </summary>
    public byte[] Bytes { get; }

    public static ParsedParts Read(byte[] canonicalBytes, IReadOnlySet<string> tokenParts)
    {
        var children = new Dictionary<string, OpenXmlElement[]>(StringComparer.Ordinal);
        if (tokenParts.Count == 0)
        {
            return new(canonicalBytes, children);
        }

        using var stream = DocxCloner.ToWritableStream(canonicalBytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc, tokenParts))
            {
                var kept = root.ChildElements.ToArray();
                foreach (var element in kept)
                {
                    LoadFully(element);
                }

                // Detached, so a kept element has no parent or sibling to lead back to the
                // document, which is closed when this returns.
                root.RemoveAllChildren();
                children[uri] = kept;
            }
        }

        return new(stream.ToArray(), children);
    }

    /// <summary>
    /// Gives <paramref name="root"/>, the emptied root of the part at <paramref name="uri"/> in a
    /// render's own package, a copy of the content that part holds in the template.
    /// </summary>
    public void Fill(string uri, OpenXmlCompositeElement root)
    {
        foreach (var element in children[uri])
        {
            root.AppendChild(element.CloneNode(true));
        }
    }

    static void LoadFully(OpenXmlElement element)
    {
        // Reaching an element is what makes its parent parse its children.
        foreach (var _ in element.Descendants())
        {
        }
    }
}
