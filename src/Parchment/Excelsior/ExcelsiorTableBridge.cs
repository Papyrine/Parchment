/// <summary>
/// Reflection adapter that constructs the closed-generic <c>Excelsior.WordTableBuilder&lt;T&gt;</c>
/// for a model element type known only at runtime, then renders the table against the host
/// <see cref="MainDocumentPart"/>.
/// </summary>
[RequiresUnreferencedCode(Unsupported.Trimming)]
[RequiresDynamicCode(Unsupported.DynamicCode)]
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

    public static Table BuildTable(Type elementType, object data, MainDocumentPart mainPart, BookmarkNames names, string? headingParagraphStyle, string? bodyParagraphStyle, string? tableStyle, Action<object>? configure)
    {
        var invoker = invokerCache.GetOrAdd(elementType, CreateInvoker);
        // Excelsior adds the relationship of a link cell to the main part itself, and the rewrite
        // below removes the ones a "#name" link no longer needs.
        RelationshipIds.Forget(mainPart);
        var table = invoker(data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure);

        // Both flows, and every table that comes back through here: a "#name" link means the same
        // thing wherever the table came from.
        ExcelsiorInternalLinks.Rewrite(table, mainPart, names);
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

        return builder.Build(mainPart);
    }

    static BuilderInvoker CreateInvoker(Type elementType)
    {
        var method = genericBuildTable.MakeGenericMethod(elementType);
        // DoNotWrapExceptions: what a Configure hook or a row getter throws reaches the render as
        // itself rather than inside a TargetInvocationException.
        return (data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure) =>
            (Table) method.Invoke(
                null,
                BindingFlags.DoNotWrapExceptions,
                null,
                [data, mainPart, headingParagraphStyle, bodyParagraphStyle, tableStyle, configure],
                null)!;
    }

    delegate Table BuilderInvoker(object data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle, string? tableStyle, Action<object>? configure);
}
