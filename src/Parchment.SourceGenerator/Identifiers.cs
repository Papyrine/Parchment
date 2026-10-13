/// <summary>
/// A name as it has to be written in generated code. <see cref="ISymbol.Name"/> is the name
/// without the <c>@</c> a keyword needs, which is the form a template uses (<c>{{ class }}</c>)
/// and the form the binding tables are keyed on, so the escape is added only where a name is
/// emitted as code. Contextual keywords are escaped too: most are legal bare, but a type cannot
/// be declared as <c>file</c>, <c>required</c> or <c>scoped</c>, and <c>@</c> is legal on any
/// identifier.
/// </summary>
static class Identifiers
{
    public static string Escape(string name)
    {
        if (SyntaxFacts.GetKeywordKind(name) == SyntaxKind.None &&
            SyntaxFacts.GetContextualKeywordKind(name) == SyntaxKind.None)
        {
            return name;
        }

        return "@" + name;
    }

    /// <summary>
    /// A type name where the type is declared. On top of the keywords, a name made of lower-case
    /// ASCII letters alone is one the compiler warns may become a keyword (CS8981) wherever it is
    /// declared bare, and a pragma in the consumer's own file does not reach the generated one.
    /// </summary>
    public static string EscapeDeclaration(string name)
    {
        foreach (var c in name)
        {
            if (c is < 'a' or > 'z')
            {
                return Escape(name);
            }
        }

        return "@" + name;
    }
}
