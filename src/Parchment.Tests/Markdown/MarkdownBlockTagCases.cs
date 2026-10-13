/// <summary>
/// Markdown templates that exercise the liquid tags beyond <c>for</c> and <c>if</c>, each with the
/// one verdict both markdown validators have to reach.
/// </summary>
/// <remarks>
/// Compiled into <c>Parchment.Tests</c> (<c>MarkdownBlockTagTests</c>, the runtime
/// <c>MarkdownReferenceValidator</c>) and linked into <c>Parchment.SourceGenerator.Tests</c>
/// (<c>MarkdownBlockTagGeneratorTests</c>, the generator's <c>MarkdownValidator</c>), so a case
/// cannot be added for one validator and not the other. Both declare the same model: a root with
/// <c>Name</c>, <c>Flag</c>, <c>Off</c>, <c>Count</c> and <c>Lines</c>, whose elements have
/// <c>Name</c> and <c>Children</c>.
/// </remarks>
static class MarkdownBlockTagCases
{
    public static IEnumerable<string> Accepted() =>
    [
        // A variable introduced inside a branch is still a variable after it.
        "{% unless Off %}{% assign n = 1 %}{% endunless %}[{{ n }}]",
        "{% unless Flag %}a{% else %}{% assign n = 1 %}{% endunless %}[{{ n }}]",
        "{% unless Off %}{% capture c %}x{% endcapture %}{% endunless %}[{{ c }}]",
        "{% case Count %}{% when 2 %}{% assign n = 1 %}{% endcase %}[{{ n }}]",
        "{% case Count %}{% when 1 %}a{% else %}{% assign n = 1 %}{% endcase %}[{{ n }}]",
        "{% case Count %}{% when 1, 2 %}{% capture c %}x{% endcapture %}{% endcase %}[{{ c }}]",

        "{% unless Name == 'x' %}{{ Name }}{% else %}{{ Count }}{% endunless %}",
        "{% case Name %}{% when Name %}{{ Count }}{% else %}{{ Name }}{% endcase %}",
        "{% capture c %}{% unless Off %}{{ Name }}{% endunless %}{% endcapture %}{{ c }}",
        "{% increment n %}{% decrement m %}",
        "{% ifchanged %}{{ Name }}{% endifchanged %}",

        "{% for l in Lines %}{% cycle 'a', 'b' %}{% endfor %}",
        "{% for l in Lines %}{% cycle Name, l.Name %}{% endfor %}",
        "{% for l in Lines %}{% cycle Name: 'a', l.Name %}{% endfor %}",

        // Loop variables stay bound through the other tags.
        "{% for l in Lines %}{% unless l.Name == Name %}{% case l.Name %}{% when Name %}{{ l.Children.size }}{% endcase %}{% endunless %}{% endfor %}",
        "{% for l in Lines %}{{ l.Name }}{% else %}{% unless Off %}{{ Name }}{% endunless %}{% endfor %}",
        "{% unless Off %}{% for l in Lines %}{% case l.Name %}{% when 'a' %}{{ forloop.index }}{% endcase %}{% endfor %}{% endunless %}",

        "<table>{% tablerow l in Lines %}{{ l.Name }}{% endtablerow %}</table>",
        "<table>{% tablerow l in Lines cols: 2 limit: 3 offset: 1 %}{{ l.Name }}{% endtablerow %}</table>",
        "<table>{% tablerow l in Lines cols: Count limit: Count offset: Count %}{{ l.Name }}{% endtablerow %}</table>",
        "<table>{% tablerow l in Lines %}{{ tablerowloop.index }}{{ tablerowloop.col }}{{ tablerowloop.row }}{% endtablerow %}</table>",
        "<table>{% tablerow l in Lines %}{% for c in l.Children %}{{ c.Name }}{{ forloop.index }}{{ tablerowloop.col }}{% endfor %}{% endtablerow %}</table>",
        "{% for l in Lines %}<table>{% tablerow c in l.Children %}{{ c.Name }}{{ l.Name }}{{ forloop.index }}{% endtablerow %}</table>{% endfor %}",
        "{% for l in Lines %}<table>{% tablerow c in l.Children cols: l.Children.size %}{{ c.Name }}{% endtablerow %}</table>{% endfor %}",
        "<table>{% tablerow l in Lines %}{% unless Off %}{% case l.Name %}{% when Name %}{{ l.Name }}{% endcase %}{% endunless %}{% endtablerow %}</table>",
        // The row variable shadows the loop variable, and the loop variable comes back after.
        "{% for l in Lines %}<table>{% tablerow l in l.Children %}{{ l.Children.size }}{% endtablerow %}</table>{{ l.Name }}{% endfor %}",
        // Nothing is known about an assign target, so neither is anything about its rows.
        "{% assign rows = Lines %}<table>{% tablerow row in rows %}{{ row.Anything }}{% endtablerow %}</table>",
        "<table>{% tablerow i in (1..Count) %}{{ i }}{% endtablerow %}</table>",
        "<table>{% tablerow l in Lines %}{% assign n = l.Name %}{% endtablerow %}</table>{{ n }}"
    ];

    /// <summary>
    /// Each template with the one path in it that names nothing.
    /// </summary>
    public static IEnumerable<(string Template, string Path)> Rejected() =>
    [
        ("{% unless Missing %}x{% endunless %}", "Missing"),
        ("{% unless Off %}{{ Missing }}{% endunless %}", "Missing"),
        ("{% unless Flag %}a{% else %}{{ Missing }}{% endunless %}", "Missing"),
        ("{% unless Off %}{% for l in Lines %}{{ l.Missing }}{% endfor %}{% endunless %}", "l.Missing"),
        ("{% unless Off %}{% assign n = Missing %}{% endunless %}", "Missing"),

        ("{% case Missing %}{% when 1 %}a{% endcase %}", "Missing"),
        ("{% case Count %}{% when Missing %}a{% endcase %}", "Missing"),
        ("{% case Count %}{% when 1, Missing %}a{% endcase %}", "Missing"),
        ("{% case Count %}{% when 2 %}{{ Missing }}{% endcase %}", "Missing"),
        ("{% case Count %}{% when 1 %}a{% when 2 %}{{ Missing }}{% endcase %}", "Missing"),
        ("{% case Count %}{% when 1 %}a{% else %}{{ Missing }}{% endcase %}", "Missing"),
        ("{% case Count %}{% when 2 %}{% for l in Lines %}{{ l.Missing }},{% endfor %}{% endcase %}", "l.Missing"),

        ("{% for l in Lines %}{% cycle Missing, 'b' %}{% endfor %}", "Missing"),
        ("{% for l in Lines %}{% cycle 'a', l.Missing %}{% endfor %}", "l.Missing"),
        ("{% cycle Missing: 'a', 'b' %}", "Missing"),

        ("{% capture c %}{% unless Off %}{{ Missing }}{% endunless %}{% endcapture %}", "Missing"),
        ("{% ifchanged %}{{ Missing }}{% endifchanged %}", "Missing"),
        ("{% for l in Lines %}x{% else %}{% unless Missing %}y{% endunless %}{% endfor %}", "Missing"),
        // A counter is not a variable: liquid keeps it apart, and the token renders blank.
        ("{% increment n %}{{ n }}", "n"),
        ("{% decrement n %}{{ n }}", "n"),

        ("<table>{% tablerow l in Missing %}x{% endtablerow %}</table>", "Missing"),
        ("<table>{% tablerow l in Lines %}{{ l.Missing }}{% endtablerow %}</table>", "l.Missing"),
        ("<table>{% tablerow l in Lines %}{{ Missing }}{% endtablerow %}</table>", "Missing"),
        ("<table>{% tablerow l in Lines cols: Missing %}x{% endtablerow %}</table>", "Missing"),
        ("<table>{% tablerow l in Lines limit: Missing %}x{% endtablerow %}</table>", "Missing"),
        ("<table>{% tablerow l in Lines offset: Missing %}x{% endtablerow %}</table>", "Missing"),
        // The arguments are read before the first row is bound.
        ("<table>{% tablerow l in Lines limit: l.Count %}x{% endtablerow %}</table>", "l.Count"),
        // Neither the row variable nor `tablerowloop` outlives the tag, and it never sets `forloop`.
        ("<table>{% tablerow l in Lines %}x{% endtablerow %}</table>{{ l.Name }}", "l.Name"),
        ("<table>{% tablerow l in Lines %}x{% endtablerow %}</table>{{ tablerowloop.index }}", "tablerowloop.index"),
        ("<table>{% tablerow l in Lines %}{{ forloop.index }}{% endtablerow %}</table>", "forloop.index"),
        ("{% for l in Lines %}{{ tablerowloop.index }}{% endfor %}", "tablerowloop.index"),
        ("<table>{% tablerow l in Lines %}{% for c in l.Children %}{{ c.Missing }}{% endfor %}{% endtablerow %}</table>", "c.Missing"),
        ("{% for l in Lines %}<table>{% tablerow c in l.Missing %}x{% endtablerow %}</table>{% endfor %}", "l.Missing")
    ];
}
