/// <summary>
/// Reflection adapter that constructs the closed-generic <c>Excelsior.WordTableBuilder&lt;T&gt;</c>
/// for a model element type known only at runtime, then renders the table against the host
/// <see cref="MainDocumentPart"/>.
/// </summary>
static class ExcelsiorTableBridge
{
    static readonly ConcurrentDictionary<Type, BuilderInvoker> invokerCache = new();

    static readonly MethodInfo genericBuildTable = typeof(ExcelsiorTableBridge)
        .GetMethods(BindingFlags.Public | BindingFlags.Static)
        .Single(_ => _ is
        {
            Name: nameof(BuildTable),
            IsGenericMethodDefinition: true
        });

    public static Table BuildTable(Type elementType, object data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle, string? tableStyle, Action<object>? configure)
    {
        var invoker = invokerCache.GetOrAdd(elementType, CreateInvoker);
        var table = invoker(data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure);

        // Both flows, and every table that comes back through here: a "#name" link means the same
        // thing wherever the table came from.
        ExcelsiorInternalLinks.Rewrite(table, mainPart);
        return table;
    }

    public static Table BuildTable<TElement>(IEnumerable<TElement> data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle, string? tableStyle, Action<object>? configure)
    {
        var builder = new WordTableBuilder<TElement>(data);
        if (headingParagraphStyle != null)
        {
            builder.HeadingParagraphStyle(headingParagraphStyle);
        }

        if (bodyParagraphStyle != null)
        {
            builder.BodyParagraphStyle(bodyParagraphStyle);
        }

        if (tableStyle != null)
        {
            builder.TableStyle(tableStyle);
        }

        // After the attribute settings, so what Configure sets wins - it is the escape hatch for
        // everything the attribute cannot say.
        configure?.Invoke(builder);

        var hadStylesPart = mainPart.StyleDefinitionsPart != null;
        var table = builder.Build(mainPart);
        if (!hadStylesPart)
        {
            PinStylesPartId(mainPart);
        }

        return table;
    }

    // Excelsior writes its built-in table styles into the styles part and, for a template or style
    // source that has none, adds the part without an id, so the SDK mints a random one. Renaming
    // afterwards rather than adding the part up front leaves it to Excelsior whether a styles part
    // is needed at all (it is not for a caller-supplied table style).
    static void PinStylesPartId(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.StyleDefinitionsPart;
        if (stylesPart == null)
        {
            return;
        }

        mainPart.ChangeIdOfPart(stylesPart, RelationshipIds.Named(mainPart, "rIdStyles"));
    }

    static BuilderInvoker CreateInvoker(Type elementType)
    {
        var method = genericBuildTable.MakeGenericMethod(elementType);
        return (data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure) =>
            (Table) method.Invoke(null, [data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure])!;
    }

    delegate Table BuilderInvoker(object data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle, string? tableStyle, Action<object>? configure);
}
