/// <summary>
/// The docx templates whose structure is wrong, with what the runtime says about each and the
/// diagnostic the source generator must raise for it at build time.
/// </summary>
/// <remarks>
/// Compiled into both test projects (linked into <c>Parchment.SourceGenerator.Tests</c>), because
/// neither can run the other's half: <c>StructuralLockstepTests</c> registers each template and
/// expects <see cref="StructuralCase.Runtime"/>, <c>StructuralLockstepGeneratorTests</c> runs the
/// generator over the same paragraphs and expects <see cref="StructuralCase.Id"/>. A rule added to
/// <c>ScopeTreeBuilder</c> or the runtime <c>TokenScanner</c> gets a case here, which fails the
/// generator half until the generator reports it too.
/// <para>
/// Every template runs against a model with <c>bool Flag</c>, <c>string Name</c> and
/// <c>Lines</c>, a list of items with a <c>Description</c>.
/// </para>
/// </remarks>
static class StructuralLockstepCases
{
    const string unbalanced = "PARCH035";
    const string unparsable = "PARCH006";
    const string sharedParagraph = "PARCH005";

    public static IEnumerable<string> Names() =>
        All.Select(_ => _.Name);

    public static StructuralCase Get(string name) =>
        All.Single(_ => _.Name == name);

    public static IReadOnlyList<StructuralCase> All { get; } =
    [
        new(
            "LoneFor",
            ["{% for line in Lines %}", "{{ line.Description }}"],
            "Missing closing tag 'endfor'.",
            unbalanced),
        new(
            "LoneIf",
            ["{% if Flag %}", "{{ Name }}"],
            "Missing closing tag 'endif'.",
            unbalanced),
        new(
            "LoneEndFor",
            ["{{ Name }}", "{% endfor %}"],
            "Unexpected '{% endfor %}' without a matching opening tag.",
            unbalanced),
        new(
            "LoneEndIf",
            ["{{ Name }}", "{% endif %}"],
            "Unexpected '{% endif %}' without a matching opening tag.",
            unbalanced),
        new(
            "ElseOutsideIf",
            ["{% else %}", "{{ Name }}"],
            "Unexpected '{% else %}' without a matching opening tag.",
            unbalanced),
        new(
            "ElsIfOutsideIf",
            ["{% elsif Flag %}", "{{ Name }}"],
            "Unexpected '{% elsif Flag %}' without a matching opening tag.",
            unbalanced),
        new(
            "ElseInsideFor",
            ["{% for line in Lines %}", "{% else %}", "{% endfor %}"],
            "Unexpected '{% else %}' without a matching opening tag.",
            unbalanced),
        new(
            "ExtraEndForAfterALoop",
            ["{% for line in Lines %}", "{{ line.Description }}", "{% endfor %}", "{% endfor %}"],
            "Unexpected '{% endfor %}' without a matching opening tag.",
            unbalanced),
        new(
            "ForCrossedWithIf",
            ["{% for line in Lines %}", "{% if Flag %}", "{{ line.Description }}", "{% endfor %}", "{% endif %}"],
            "Unexpected '{% endfor %}' without a matching opening tag.",
            unbalanced),
        new(
            "IfCrossedWithFor",
            ["{% if Flag %}", "{% for line in Lines %}", "{{ line.Description }}", "{% endif %}", "{% endfor %}"],
            "Unexpected '{% endif %}' without a matching opening tag.",
            unbalanced),
        new(
            "IfLeftOpenInsideFor",
            ["{% for line in Lines %}", "{% if Flag %}", "{{ line.Description }}", "{% endfor %}"],
            "Unexpected '{% endfor %}' without a matching opening tag.",
            unbalanced),
        new(
            "ForLeftOpenInsideElse",
            ["{% if Flag %}", "A", "{% else %}", "{% for line in Lines %}", "{% endif %}"],
            "Unexpected '{% endif %}' without a matching opening tag.",
            unbalanced),
        new(
            "NestedIfLeftOpen",
            ["{% if Flag %}", "{% if Flag %}", "{{ Name }}", "{% endif %}"],
            "Missing closing tag 'endif'.",
            unbalanced),
        new(
            "ElseAfterElse",
            ["{% if Flag %}", "A", "{% else %}", "B", "{% else %}", "C", "{% endif %}"],
            "Unexpected '{% else %}' after '{% else %}'.",
            unbalanced),
        new(
            "ElsIfAfterElse",
            ["{% if Flag %}", "A", "{% else %}", "B", "{% elsif Flag %}", "C", "{% endif %}"],
            "Unexpected '{% elsif Flag %}' after '{% else %}'.",
            unbalanced),
        new(
            "EmptyToken",
            ["{{ }}"],
            "Failed to parse liquid token: ",
            unparsable),
        new(
            "MethodCallSyntax",
            ["{{ Name.ToUpper() }}"],
            "Failed to parse liquid token: ",
            unparsable),
        new(
            "FilterWithoutAName",
            ["{{ Name | }}"],
            "Failed to parse liquid token: ",
            unparsable),
        new(
            "TwoBlockTagsInOneParagraph",
            ["{% if Flag %}{% endif %}"],
            "A paragraph contains more than one block tag. Place each block tag in its own paragraph.",
            sharedParagraph),
        new(
            "BlockTagBesideASubstitution",
            ["{% if Flag %}{{ Name }}", "{% endif %}"],
            "A paragraph mixes block tags with substitution tokens. Block tags must sit in their own paragraphs.",
            sharedParagraph),
        new(
            "BlockTagBesideText",
            ["Lines: {% for line in Lines %}", "{{ line.Description }}", "{% endfor %}"],
            "A paragraph contains a block tag alongside non-token text. Block tags must sit in their own paragraphs.",
            sharedParagraph)
    ];
}

/// <param name="Runtime">
/// What the <c>ParchmentRegistrationException</c> says. For <c>PARCH035</c> and <c>PARCH006</c>
/// the diagnostic carries the same text; <c>PARCH005</c> predates this and has wording of its own.
/// </param>
sealed record StructuralCase(string Name, string[] Paragraphs, string Runtime, string Id);
