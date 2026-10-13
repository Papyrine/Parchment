using Microsoft.CodeAnalysis.CSharp.Syntax;

/// <summary>
/// Attribute arguments are constants the compiler has already accepted, which is a wider set than
/// the values that mean something: any integer casts to an enum, and a string can hold anything.
/// Whatever arrives has to come out as a diagnostic on the attribute or as code that compiles —
/// never as a syntax error in a generated file.
/// </summary>
public class AttributeArgumentGeneratorTests
{
    static string ModelWithProtection(string argument) =>
        $$"""
          using Parchment;

          namespace Sample;

          [ParchmentModel(Protection = {{argument}})]
          public partial class Letter
          {
              public string Name { get; set; } = "";
          }
          """;

    [Test]
    [Arguments("(ProtectionMode)5", 5)]
    [Arguments("(ProtectionMode)(-1)", -1)]
    [Arguments("(ProtectionMode)int.MaxValue", int.MaxValue)]
    [Arguments("ProtectionMode.None | (ProtectionMode)2", 3)]
    public async Task UndefinedProtection_ReportsParch036_Docx(string argument, int value)
    {
        var source = ModelWithProtection(argument);
        var result = GeneratorDriver.Run(source, "{{ Name }}");

        var diagnostic = result.Results.Single().Diagnostics.Single();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH036");
        await Assert.That(diagnostic.Severity).IsEqualTo(DiagnosticSeverity.Error);
        await Assert.That(diagnostic.GetMessage()).Contains("'Letter'");
        await Assert.That(diagnostic.GetMessage()).Contains($" {value},");
        // On the attribute, where the value was written.
        var span = diagnostic.Location.SourceSpan;
        await Assert.That(source.Substring(span.Start, span.Length)).StartsWith("ParchmentModel(Protection");

        // What is emitted beside the error still parses, as the cast the author wrote.
        var tree = result.GeneratedTrees.Single();
        await Assert.That(tree.GetDiagnostics()).IsEmpty();
        await Assert.That(tree.ToString()).Contains($", (global::Parchment.ProtectionMode)({value}));");
    }

    // Protection is ignored by a markdown template, but an undefined value is no more meaningful
    // there, and the model may be given a docx template tomorrow.
    [Test]
    public async Task UndefinedProtection_ReportsParch036_Markdown()
    {
        var result = GeneratorDriver.RunMarkdown(ModelWithProtection("(ProtectionMode)5"), "Hi {{ Name }}");

        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEquivalentTo(["PARCH036"]);
        await Assert.That(result.GeneratedTrees.Single().GetDiagnostics()).IsEmpty();
    }

    // The template diagnostics are still worth having in the same build.
    [Test]
    public async Task UndefinedProtection_DoesNotHideTemplateDiagnostics()
    {
        var result = GeneratorDriver.Run(ModelWithProtection("(ProtectionMode)5"), "{{ Missing }}");

        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEquivalentTo(["PARCH001", "PARCH036"]);
    }

    // Defined values, however they are spelled, are written by name.
    [Test]
    [Arguments("ProtectionMode.None", ", global::Parchment.ProtectionMode.None);")]
    [Arguments("(ProtectionMode)1", ", global::Parchment.ProtectionMode.None);")]
    [Arguments("ProtectionMode.None | ProtectionMode.WhenEditable", ", global::Parchment.ProtectionMode.None);")]
    [Arguments("ProtectionMode.WhenEditable", "\"));")]
    [Arguments("(ProtectionMode)0", "\"));")]
    [Arguments("default", "\"));")]
    public async Task DefinedProtection_EmitsByName(string argument, string expectedTail)
    {
        var result = GeneratorDriver.Run(ModelWithProtection(argument), "{{ Name }}");

        await Assert.That(result.Results.Single().Diagnostics).IsEmpty();
        var tree = result.GeneratedTrees.Single();
        await Assert.That(tree.GetDiagnostics()).IsEmpty();
        await Assert.That(tree.ToString()).Contains(expectedTail);
    }

    // Quotes, backslashes, braces, line breaks, a NUL, the Unicode line separator:
    // each would end or corrupt a string literal written out verbatim.
    const string awkward = "a\"b\\c{0}}\r\n\t\0\u2028'";
    const string awkwardSource =
        """
        "a\"b\\c{0}}\r\n\t\0\u2028'"
        """;

    [Test]
    public async Task StringArguments_RoundTripThroughTheGeneratedLiteral()
    {
        var source =
            $$"""
              using System;
              using System.Collections.Generic;
              using Parchment;

              namespace Sample;

              public class Line
              {
                  public string Description { get; set; } = "";
              }

              [ParchmentModel]
              public partial class Invoice
              {
                  [ExcelsiorTable(HeadingParagraphStyle = {{awkwardSource}}, BodyParagraphStyle = "", TableStyle = null)]
                  public List<Line> Lines { get; set; } = new();

                  [EditableField(DateFormat = {{awkwardSource}})]
                  public DateTime Due { get; set; }

                  [EditableField(DateFormat = "")]
                  public DateTime Paid { get; set; }

                  [EditableField(DateFormat = null, MultiLine = false)]
                  public string Notes { get; set; } = "";
              }
              """;
        var result = GeneratorDriver.Run(source, "{{ Lines }}", "{{ Due }}", "{{ Paid }}", "{{ Notes }}");

        await Assert.That(result.Results.Single().Diagnostics).IsEmpty();
        var tree = result.GeneratedTrees.Single();
        await Assert.That(tree.GetDiagnostics()).IsEmpty();

        // The table entry: heading style, then the empty body style, then the absent table style.
        await Assert.That(StringArguments(tree, "Lines")).IsEqualTo(Join(awkward, "", null));
        await Assert.That(StringArguments(tree, "Due")).IsEqualTo(Join(awkward));
        await Assert.That(StringArguments(tree, "Paid")).IsEqualTo(Join(""));
        await Assert.That(StringArguments(tree, "Notes")).IsEqualTo("<null>");
        // One physical line per entry is what the indent pass relies on; an unescaped line break
        // inside a literal would split it.
        var lines = tree.ToString().Split('\n');
        await Assert.That(lines.Count(_ => _.Contains("new(\"Due\", global::Parchment.Generated"))).IsEqualTo(1);
    }

    // The string and null literals that follow the key in a map entry's `new("key", ...)` argument
    // list, as the values the compiler reads back out of them. The accessor entries are the
    // two-argument ones; the style and date-format arguments are the only literals in the others.
    static string StringArguments(SyntaxTree tree, string key)
    {
        var creation = tree.GetRoot()
            .DescendantNodes()
            .OfType<ImplicitObjectCreationExpressionSyntax>()
            .Single(_ => _.ArgumentList.Arguments.Count > 2 &&
                         _.ArgumentList.Arguments[0].ToString() == $"\"{key}\"");
        return Join(creation.ArgumentList.Arguments
            .Skip(1)
            .Select(_ => _.Expression)
            .OfType<LiteralExpressionSyntax>()
            .Where(_ => _.IsKind(SyntaxKind.StringLiteralExpression) ||
                        _.IsKind(SyntaxKind.NullLiteralExpression))
            .Select(_ => (string?)_.Token.Value)
            .ToArray());
    }

    static string Join(params string?[] values) =>
        string.Join(" | ", values.Select(_ => _ ?? "<null>"));
}
