public class MarkdownValidatorTests
{
    const string letterModel =
        """
        using Parchment;

        namespace Sample;

        public class Customer
        {
            public string Name { get; set; } = "";
            public string Region { get; set; } = "";
        }

        [ParchmentModel("template.md")]
        public partial class Letter
        {
            public Customer Customer { get; set; } = new();
        }
        """;

    const string invoiceModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
            public int Quantity { get; set; }
        }

        [ParchmentModel("template.md")]
        public partial class Invoice
        {
            public List<Line> Lines { get; set; } = new();
            public List<Line> Adjustments { get; set; } = new();
        }
        """;

    [Test]
    public async Task NestedLoopsShadowSameVariable()
    {
        // Inner `line` shadows outer `line`. Inner body's `line.Quantity` resolves against the
        // inner Line; after the inner endfor, the outer `line` is restored and `line.Description`
        // resolves against the outer Line. No PARCH001 expected.
        var result = GeneratorDriver.RunMarkdown(
            invoiceModel,
            """
            {% for line in Lines %}
              {% for line in Adjustments %}
                {{ line.Quantity }}
              {% endfor %}
              {{ line.Description }}
            {% endfor %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task LoopVariableLeaksAfterEndfor()
    {
        // After `{% endfor %}`, `line` is no longer in scope, so `line.Description` must trip
        // PARCH001 — guards against the WalkFor cleanup forgetting to scope.Remove the loop var.
        var result = GeneratorDriver.RunMarkdown(
            invoiceModel,
            """
            {% for line in Lines %}
              {{ line.Description }}
            {% endfor %}
            {{ line.Description }}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        var missing = diagnostics.Where(_ => _.Id == "PARCH001").ToList();
        await Assert.That(missing.Count).IsEqualTo(1);
        await Assert.That(missing[0].GetMessage()).Contains("line.Description");
    }

    [Test]
    public async Task ForElseBranchValidates()
    {
        // Bad reference inside `{% else %}` branch of a for loop must still trip PARCH001.
        var result = GeneratorDriver.RunMarkdown(
            invoiceModel,
            """
            {% for line in Lines %}
              {{ line.Description }}
            {% else %}
              {{ Lines.Bad }}
            {% endfor %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        var missing = diagnostics.Where(_ => _.Id == "PARCH001").ToList();
        await Assert.That(missing.Count).IsEqualTo(1);
        await Assert.That(missing[0].GetMessage()).Contains("Lines.Bad");
    }

    [Test]
    public async Task IfConditionAndBodyValidate()
    {
        // Both the condition expression and the body must be walked — two PARCH001s expected.
        var result = GeneratorDriver.RunMarkdown(
            letterModel,
            """
            {% if Customer.BadCondition %}
              {{ Customer.BadBody }}
            {% endif %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        var missing = diagnostics.Where(_ => _.Id == "PARCH001").ToList();
        await Assert.That(missing.Count).IsEqualTo(2);
    }

    [Test]
    public async Task ElsifAndElseBranchesValidate()
    {
        var result = GeneratorDriver.RunMarkdown(
            letterModel,
            """
            {% if Customer.Name %}
              ok
            {% elsif Customer.BadElsif %}
              {{ Customer.BadElsifBody }}
            {% else %}
              {{ Customer.BadElseBody }}
            {% endif %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        var missing = diagnostics.Where(_ => _.Id == "PARCH001").ToList();
        await Assert.That(missing.Count).IsEqualTo(3);
    }

    [Test]
    public async Task RangeLoopSourceParsesAndBodyWalks()
    {
        // `(1..3)` is a RangeExpression — TryGetMemberPath returns null, so no PARCH002 should
        // fire (we can't validate the source as a member). The body still walks: `Customer.Name`
        // resolves cleanly against the root model. Confirms the non-member source path doesn't
        // throw or short-circuit the body.
        var result = GeneratorDriver.RunMarkdown(
            letterModel,
            """
            {% for i in (1..3) %}
              {{ Customer.Name }}
            {% endfor %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task BinaryConditionExpressionsAreCollected()
    {
        // `Customer.A == Customer.B` must produce two PARCH001s (A and B both missing) —
        // confirms ExpressionPathCollector descends into binary expressions via the default
        // AstVisitor traversal.
        var result = GeneratorDriver.RunMarkdown(
            letterModel,
            """
            {% if Customer.BadA == Customer.BadB %}
              ok
            {% endif %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        var missing = diagnostics.Where(_ => _.Id == "PARCH001").ToList();
        await Assert.That(missing.Count).IsEqualTo(2);
    }

    [Test]
    public async Task LoopSource_IndexerWithStringLiteral_ResolvesElementType()
    {
        // `Customer['Lines']` is Fluid-equivalent to `Customer.Lines`. The loop body's `line.Quantity`
        // must validate against Line (the element type), not against the root Invoice — otherwise
        // every body access produces a false PARCH001.
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Description { get; set; } = "";
                public int Quantity { get; set; }
            }

            public class Owner
            {
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentModel("template.md")]
            public partial class Invoice
            {
                public Owner Customer { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.RunMarkdown(
            source,
            """
            {% for line in Customer['Lines'] %}
              {{ line.Quantity }}
              {{ line.Description }}
            {% endfor %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task BadLoopSourceCascadeIsBounded()
    {
        // The loop source `Customer` resolves but isn't enumerable → PARCH002. The body's
        // `line.Name` accesses Customer.Name (root-fallback binding) and the root Letter has
        // no `Name` member, so we expect PARCH001. The point of this test: the body walk runs
        // (no crash, no early return) and the cascade is bounded — accessing real members on
        // the root type does NOT trigger extra noise.
        var result = GeneratorDriver.RunMarkdown(
            letterModel,
            """
            {% for line in Customer %}
              {{ line.Customer.Name }}
            {% endfor %}
            """);
        var diagnostics = result.Results.Single().Diagnostics;
        // PARCH002 fires once for the non-enumerable source.
        await Assert.That(diagnostics.Count(_ => _.Id == "PARCH002")).IsEqualTo(1);
        // Body access on a real path through the root (line → Letter, .Customer.Name) resolves
        // cleanly — no PARCH001 cascade.
        await Assert.That(diagnostics.Count(_ => _.Id == "PARCH001")).IsEqualTo(0);
    }
}
