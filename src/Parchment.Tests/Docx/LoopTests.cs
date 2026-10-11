// ReSharper disable PartialTypeWithSinglePart
public partial class LoopTests
{
    [Test]
    public async Task ParagraphScopeLoop()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Items:

            {% for line in Lines %}

            - {{ line.Description }}: {{ line.Quantity }} x {{ line.UnitPrice }}

            {% endfor %}

            End.
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task NestedLoop()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for group in Groups %}

            {{ group.Name }}

            {% for item in group.Items %}

            - {{ item }}

            {% endfor %}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<NestedModel>(template);
        using var stream = new MemoryStream();
        await store.Render(
            new NestedModel
            {
                Groups =
                [
                    new()
                    {
                        Name = "Fruit",
                        Items = ["apple", "pear"]
                    },
                    new()
                    {
                        Name = "Tools",
                        Items = ["hammer", "saw", "drill"]
                    }
                ]
            },
            stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class NestedModel
    {
        public required IReadOnlyList<NestedGroup> Groups { get; init; }
    }

    public class NestedGroup
    {
        public required string Name { get; init; }
        public required IReadOnlyList<string> Items { get; init; }
    }

    [Test]
    public async Task EmptyLoop_ProducesNoIterations()
    {
        // Zero-iteration loop must remove the body without leaving the open/close anchor
        // paragraphs behind. Surrounding paragraphs ("Before"/"After") stay put.
        using var template = DocxTemplateBuilder.Build(
            """
            Before

            {% for line in Lines %}

            - {{ line.Description }}

            {% endfor %}

            After
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);

        var invoice = SampleData.Invoice();
        var emptied = new Invoice
        {
            Number = invoice.Number,
            IssueDate = invoice.IssueDate,
            DueDate = invoice.DueDate,
            Customer = invoice.Customer,
            Lines = [],
            Currency = invoice.Currency,
            Notes = invoice.Notes,
            Tags = invoice.Tags
        };

        using var stream = new MemoryStream();
        await store.Render(emptied, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class TripleNestModel
    {
        public required IReadOnlyList<Level1> Sections { get; init; }
    }

    public class Level1
    {
        public required string Title { get; init; }
        public required IReadOnlyList<Level2> Subsections { get; init; }
    }

    public class Level2
    {
        public required string Heading { get; init; }
        public required IReadOnlyList<string> Bullets { get; init; }
    }

    [Test]
    public async Task ThreeLevelNestedLoops()
    {
        // ProcessLoopAsync clones bodies onto a scratch Body before recursing — three levels deep
        // exercises that recursion path and the per-iteration anchor refresh more than the existing
        // two-level NestedLoop test.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for s in Sections %}

            # {{ s.Title }}

            {% for sub in s.Subsections %}

            ## {{ sub.Heading }}

            {% for b in sub.Bullets %}

            - {{ b }}

            {% endfor %}

            {% endfor %}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TripleNestModel>(template);

        var model = new TripleNestModel
        {
            Sections =
            [
                new()
                {
                    Title = "Alpha",
                    Subsections =
                    [
                        new()
                        {
                            Heading = "A.1",
                            Bullets = ["a", "b"]
                        },
                        new()
                        {
                            Heading = "A.2",
                            Bullets = ["c"]
                        }
                    ]
                },
                new()
                {
                    Title = "Beta",
                    Subsections =
                    [
                        new()
                        {
                            Heading = "B.1",
                            Bullets = ["x", "y", "z"]
                        }
                    ]
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class LeakModel
    {
        public required string Item { get; init; }
        public required IReadOnlyList<string> Items { get; init; }
    }

    [Test]
    public async Task LoopVariable_DoesNotLeak_AfterEndfor()
    {
        // `Item` is the loop variable AND a model property of the same name. Inside the loop,
        // `{{ Item }}` is the iteration value (shadows the model property). After `{% endfor %}`,
        // `{{ Item }}` must resolve to the model's `Item` — NOT to the last iteration's value.
        // ScopeTreeRunner.ProcessLoopAsync used to set on the root scope directly, leaking the
        // loop binding past `{% endfor %}` so it permanently shadowed the model property.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for Item in Items %}

            - {{ Item }}

            {% endfor %}

            After:{{ Item }}.
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<LeakModel>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new LeakModel { Item = "model-property", Items = ["alpha", "beta"] },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();

        await Assert.That(paragraphs).Contains("- alpha");
        await Assert.That(paragraphs).Contains("- beta");
        await Assert.That(paragraphs).Contains("After:model-property.");
        await Assert.That(paragraphs).DoesNotContain("After:beta.");
    }

    [ParchmentBindable]
    public partial class NestedShadowModel
    {
        public required IReadOnlyList<OuterItem> Outers { get; init; }
    }

    public class OuterItem
    {
        public required string OuterName { get; init; }
        public required IReadOnlyList<InnerItem> Inners { get; init; }
    }

    public class InnerItem
    {
        public required string InnerName { get; init; }
    }

    [Test]
    public async Task LoopVariable_SameName_NestedLoops_OuterRestoredAfterInner()
    {
        // Outer loop binds `x` to OuterItem. Inner loop with the same name `x` binds to InnerItem.
        // After the inner `{% endfor %}`, the outer's `{{ x.OuterName }}` must resolve against the
        // OUTER item, not stay bound to the inner's last value.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for x in Outers %}

            outer={{ x.OuterName }}

            {% for x in x.Inners %}

            inner={{ x.InnerName }}

            {% endfor %}

            after-inner={{ x.OuterName }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<NestedShadowModel>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new NestedShadowModel
            {
                Outers =
                [
                    new()
                    {
                        OuterName = "A",
                        Inners = [new() { InnerName = "a1" }, new() { InnerName = "a2" }]
                    },
                    new()
                    {
                        OuterName = "B",
                        Inners = [new() { InnerName = "b1" }]
                    }
                ]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();

        await Assert.That(paragraphs).Contains("outer=A");
        await Assert.That(paragraphs).Contains("inner=a1");
        await Assert.That(paragraphs).Contains("inner=a2");
        await Assert.That(paragraphs).Contains("after-inner=A");
        await Assert.That(paragraphs).Contains("outer=B");
        await Assert.That(paragraphs).Contains("inner=b1");
        await Assert.That(paragraphs).Contains("after-inner=B");
    }

    [ParchmentBindable]
    public partial class RangeModel
    {
        public required int Count { get; init; }
        public required IReadOnlyList<InnerItem> Inners { get; init; }
    }

    static async Task<List<string>> RenderRange(string markdown, int count = 0)
    {
        using var template = DocxTemplateBuilder.Build(markdown);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<RangeModel>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new RangeModel
            {
                Count = count,
                Inners = [new() { InnerName = "a" }, new() { InnerName = "b" }]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();
    }

    static void RegisterRange(string markdown)
    {
        using var template = DocxTemplateBuilder.Build(markdown);
        new TemplateStore().RegisterDocxTemplate<RangeModel>(template);
    }

    // A range has no member path, so nothing about it resolves against the model. That is an
    // unknown, not a mistake: registration used to reject it as "Loop source '' does not resolve to
    // an enumerable type", while the markdown flow accepted the same loop.
    [Test]
    public async Task RangeLoop_LiteralBounds_Renders()
    {
        var paragraphs = await RenderRange(
            """
            {% for i in (1..3) %}

            Row {{ i }}

            {% endfor %}

            End.
            """);

        await Assert.That(paragraphs).IsEquivalentTo(["Row 1", "Row 2", "Row 3", "End."]);
    }

    [Test]
    public async Task RangeLoop_MemberBound_Renders()
    {
        var paragraphs = await RenderRange(
            """
            {% for i in (1..Count) %}

            Row {{ i }}

            {% endfor %}
            """,
            count: 2);

        await Assert.That(paragraphs).IsEquivalentTo(["Row 1", "Row 2"]);
    }

    // A member used as a bound is a reference like any other, so a typo in one still fails.
    [Test]
    public async Task RangeLoop_UnknownMemberBound_FailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => RegisterRange(
                """
                {% for i in (1..NoSuchMember) %}

                Row {{ i }}

                {% endfor %}
                """));
        await Assert.That(exception.Message).Contains("NoSuchMember");
    }

    // The loop variable is untyped: it exists, but members reached through it are not checked.
    [Test]
    public void RangeLoop_LoopVariableMembersAreNotChecked() =>
        RegisterRange(
            """
            {% for i in (1..3) %}

            {{ i.Anything }}

            {% if i.Anything %}

            x

            {% endif %}

            {% for inner in i.Anything %}

            {{ inner.Anything }}

            {% endfor %}

            {% endfor %}
            """);

    [Test]
    public async Task RangeLoop_VariableDoesNotLeak_AfterEndfor()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => RegisterRange(
                """
                {% for i in (1..3) %}

                Row {{ i }}

                {% endfor %}

                {{ i }}
                """));
        await Assert.That(exception.Message).Contains("'i'");
    }

    // A typed loop nested in a range loop is still checked, and still renders per iteration.
    [Test]
    public async Task RangeLoop_NestedTypedLoop_StaysValidated()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => RegisterRange(
                """
                {% for i in (1..2) %}

                {% for inner in Inners %}

                {{ inner.NoSuchMember }}

                {% endfor %}

                {% endfor %}
                """));
        await Assert.That(exception.Message).Contains("NoSuchMember");

        var paragraphs = await RenderRange(
            """
            {% for i in (1..2) %}

            {% for inner in Inners %}

            {{ i }}{{ inner.InnerName }}

            {% endfor %}

            {% endfor %}
            """);

        await Assert.That(paragraphs).IsEquivalentTo(["1a", "1b", "2a", "2b"]);
    }

    // Matched by name, the way a consumer declares it.
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    [ParchmentBindable]
    public partial class ShadowedHtmlModel
    {
        public required HtmlPerson Customer { get; init; }
        public required List<HtmlPerson> Customers { get; init; }
    }

    public class HtmlPerson
    {
        [Html]
        public required string Notes { get; init; }
    }

    static ShadowedHtmlModel ShadowedHtml() =>
        new()
        {
            Customer = new()
            {
                Notes = "<b>ROOT</b>"
            },
            Customers =
            [
                new()
                {
                    Notes = "<i>one</i>"
                },
                new()
                {
                    Notes = "two"
                }
            ]
        };

    // Root-keyed dispatch matches a token by its dotted path, and `customer.Notes` reads the same
    // as the root's `Customer.Notes`. Inside the loop it is the item's member: the root value must
    // not be rendered in its place, once per item.
    [Test]
    public async Task LoopVariableNamedAfterRootHtmlMemberReadsTheItem()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Customer.Notes }}

            {% for customer in Customers %}

            {{ customer.Notes }}

            {% endfor %}

            {{ Customer.Notes }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedHtmlModel>(template);

        var paragraphs = await RenderParagraphs(store, ShadowedHtml());

        // Outside the loop the root member is still converted as html.
        await Assert.That(paragraphs).IsEquivalentTo(["ROOT", "<i>one</i>", "two", "ROOT"]);
    }

    // Fluid finds a loop variable by its exact spelling and a model member without regard to
    // case, so the root's own spelling still reaches the root while the loop is open.
    [Test]
    public async Task LoopVariableShadowsOnlyItsOwnSpelling()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for customer in Customers %}

            {{ Customer.Notes }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedHtmlModel>(template);

        var paragraphs = await RenderParagraphs(store, ShadowedHtml());

        await Assert.That(paragraphs).IsEquivalentTo(["ROOT", "ROOT"]);
    }

    // The plain-member-access rule belongs to the root member. The item's member goes through
    // Fluid, where a filter is applied like any other.
    [Test]
    public async Task FilterOnLoopVariableNamedAfterRootHtmlMemberIsAccepted()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for customer in Customers %}

            {{ customer.Notes | upcase }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedHtmlModel>(template);

        var paragraphs = await RenderParagraphs(store, ShadowedHtml());

        await Assert.That(paragraphs).IsEquivalentTo(["<I>ONE</I>", "TWO"]);
    }

    [ParchmentBindable]
    public partial class ShadowedTableModel
    {
        public required TableOrder Order { get; init; }
        public required List<OrderSummary> Orders { get; init; }
    }

    public class TableOrder
    {
        [ExcelsiorTable]
        public required List<TableLine> Lines { get; init; }
    }

    public class TableLine
    {
        public required string Description { get; init; }
    }

    public class OrderSummary
    {
        public required int Lines { get; init; }
    }

    [Test]
    public async Task LoopVariableNamedAfterRootExcelsiorTableReadsTheItem()
    {
        // The second token would fail both [ExcelsiorTable] rules, shared paragraph and filter,
        // were it taken for the root's table.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for order in Orders %}

            {{ order.Lines }}

            Next: {{ order.Lines | plus: 1 }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedTableModel>(template);

        var model = new ShadowedTableModel
        {
            Order = new()
            {
                Lines =
                [
                    new()
                    {
                        Description = "ROOT"
                    }
                ]
            },
            Orders =
            [
                new()
                {
                    Lines = 3
                },
                new()
                {
                    Lines = 5
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(Texts(body)).IsEquivalentTo(["3", "Next: 4", "5", "Next: 6"]);
    }

    [ParchmentBindable]
    public partial class ShadowedListModel
    {
        public required TaggedItem Item { get; init; }
        public required List<LabelledItem> Items { get; init; }
    }

    public class TaggedItem
    {
        public required List<string> Tags { get; init; }
    }

    public class LabelledItem
    {
        public required string Tags { get; init; }
    }

    [Test]
    public async Task LoopVariableNamedAfterRootStringListReadsTheItem()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for item in Items %}

            {{ item.Tags }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedListModel>(template);

        var model = new ShadowedListModel
        {
            Item = new()
            {
                Tags = ["root-1", "root-2"]
            },
            Items =
            [
                new()
                {
                    Tags = "one"
                },
                new()
                {
                    Tags = "two"
                }
            ]
        };

        var paragraphs = await RenderParagraphs(store, model);

        await Assert.That(paragraphs).IsEquivalentTo(["one", "two"]);
    }

    [ParchmentBindable]
    public partial class ShadowedEditableModel
    {
        public required EditablePerson Customer { get; init; }
        public required List<PlainPerson> Customers { get; init; }
    }

    public class EditablePerson
    {
        [EditableField]
        public string Notes { get; set; } = "";
    }

    public class PlainPerson
    {
        public required string Notes { get; init; }
    }

    // Taken for the root's editable member, the loop's tokens are a second and third use of its
    // path, and one of them filtered: registration refused the template on both counts.
    [Test]
    public async Task LoopVariableNamedAfterRootEditableMemberReadsTheItem()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Customer.Notes }}

            {% for customer in Customers %}

            {{ customer.Notes }} and {{ customer.Notes | upcase }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowedEditableModel>(template);

        var model = new ShadowedEditableModel
        {
            Customer = new()
            {
                Notes = "ROOT"
            },
            Customers =
            [
                new()
                {
                    Notes = "one"
                },
                new()
                {
                    Notes = "two"
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var controls = body.Descendants<SdtElement>().ToList();

        await Assert.That(controls.Count).IsEqualTo(1);
        await Assert.That(controls[0].InnerText).IsEqualTo("ROOT");
        await Assert.That(Texts(body)).IsEquivalentTo(["ROOT", "one and ONE", "two and TWO"]);
    }

    static async Task<List<string>> RenderParagraphs<TModel>(TemplateStore store, TModel model)
        where TModel : class
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return Texts(doc.MainDocumentPart!.Document!.Body!);
    }

    static List<string> Texts(Body body) =>
        body.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();
}
