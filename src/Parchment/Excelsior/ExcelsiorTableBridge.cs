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

    public static Table BuildTable(Type elementType, object data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle)
    {
        var invoker = invokerCache.GetOrAdd(elementType, CreateInvoker);
        return invoker(data, mainPart, headingParagraphStyle, bodyParagraphStyle);
    }

    public static Table BuildTable<TElement>(IEnumerable<TElement> data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle)
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

        return builder.Build(mainPart);
    }

    static BuilderInvoker CreateInvoker(Type elementType)
    {
        var method = genericBuildTable.MakeGenericMethod(elementType);
        return (data, mainPart, headingParagraphStyle, bodyParagraphStyle) =>
            (Table) method.Invoke(null, [data, mainPart, headingParagraphStyle, bodyParagraphStyle])!;
    }

    delegate Table BuilderInvoker(object data, MainDocumentPart mainPart, string? headingParagraphStyle, string? bodyParagraphStyle);
}
