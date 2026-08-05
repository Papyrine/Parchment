/// <summary>
/// One <c>WordNumberingState</c> is allocated per render and shared across every code path that
/// can introduce list numbering — string-list auto-bullets, [Markdown] content, loop bodies, and
/// if-branches. These tests assert that invariant on every path: a render that produces N
/// equivalent lists must produce exactly one bullet (or ordered) abstract definition shared by N
/// numbering instances.
///
/// Before the shared-state fix, each token built its own <c>WordNumberingState</c>, allocating
/// duplicate abstracts whose final IDs depended on OOXML SDK save-time reordering — making rendered
/// docx bytes non-deterministic across runs.
/// </summary>
public class NumberingStateSharingTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    public class TwoMarkdownDoc
    {
        [Markdown]
        public required string FirstBody { get; init; }

        [Markdown]
        public required string SecondBody { get; init; }
    }

    public class MixedDoc
    {
        public required IReadOnlyList<string> Tags { get; init; }

        [Markdown]
        public required string Body { get; init; }
    }

    public class GroupsDoc
    {
        public required IReadOnlyList<Group> Groups { get; init; }
        public required IReadOnlyList<string> Outer { get; init; }
    }

    public class Group
    {
        public required IReadOnlyList<string> Items { get; init; }
    }

    public class IfDoc
    {
        public required bool ShowInner { get; init; }
        public required IReadOnlyList<string> Inner { get; init; }
        public required IReadOnlyList<string> Outer { get; init; }
    }

    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    public class TwoHtmlDoc
    {
        [Html]
        public required string FirstBody { get; init; }

        [Html]
        public required string SecondBody { get; init; }
    }

    public class HtmlIfDoc
    {
        public required bool ShowInner { get; init; }

        [Html]
        public required string Inner { get; init; }

        [Html]
        public required string Outer { get; init; }
    }

    public class HtmlAndMarkdownDoc
    {
        [Html]
        public required string HtmlBody { get; init; }

        [Markdown]
        public required string MdBody { get; init; }
    }

    [Test]
    public async Task TwoHtmlPropertiesShareOneBulletAbstract()
    {
        // Two [Html] properties each containing a <ul> go through OpenXmlHtml.WordHtmlConverter;
        // both must reuse the shared HtmlNumberingSession's bullet abstract rather than each
        // allocating a fresh one.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ FirstBody }}

            {{ SecondBody }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TwoHtmlDoc>(template);

        var model = new TwoHtmlDoc
        {
            FirstBody = "<ul><li>a1</li><li>a2</li></ul>",
            SecondBody = "<ul><li>b1</li><li>b2</li></ul>"
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 2);
    }

    [Test]
    public async Task HtmlInIfBranchSharesBulletAbstract()
    {
        // The inner ScopeTreeRunner spawned for {% if %} inherits the parent's WordNumberingState,
        // which carries the HtmlNumberingSession. An [Html] token inside the branch and one outside
        // must share the same bullet abstract.
        using var template = DocxTemplateBuilder.Build(
            """
            {% if ShowInner %}

            {{ Inner }}

            {% endif %}

            {{ Outer }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlIfDoc>(template);

        var model = new HtmlIfDoc
        {
            ShowInner = true,
            Inner = "<ul><li>i1</li><li>i2</li></ul>",
            Outer = "<ul><li>o1</li><li>o2</li></ul>"
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 2);
    }

    [Test]
    public async Task HtmlAndMarkdownUseSeparateAbstracts()
    {
        // [Html] and [Markdown] go through different rendering stacks: OpenXmlHtml uses its own
        // 9-level bullet abstract, Parchment's WordNumberingState creates a separate 3-level one.
        // The two abstracts must NOT be collapsed — each list type keeps its own definition.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ HtmlBody }}

            {{ MdBody }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlAndMarkdownDoc>(template);

        var model = new HtmlAndMarkdownDoc
        {
            HtmlBody = "<ul><li>h1</li><li>h2</li></ul>",
            MdBody = "- m1\n- m2"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var bulletAbstracts = numbering
            .Elements<AbstractNum>()
            .Where(a =>
                a.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0) is { } l &&
                l.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
            .ToList();

        // One from OpenXmlHtml (9 levels), one from WordNumberingState (3 levels) — distinct.
        await Assert.That(bulletAbstracts.Count).IsEqualTo(2);
        await Assert.That(numbering.Elements<NumberingInstance>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task TwoMarkdownPropertiesShareOneBulletAbstract()
    {
        // Two [Markdown] properties each containing a bullet list go through
        // OpenXmlMarkdownRenderer; both must reuse the shared WordNumberingState's bullet abstract
        // rather than each constructing a fresh state.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ FirstBody }}

            {{ SecondBody }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TwoMarkdownDoc>(template);

        var model = new TwoMarkdownDoc
        {
            FirstBody = "- a1\n- a2",
            SecondBody = "- b1\n- b2"
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 2);
    }

    [Test]
    public async Task StringListAndMarkdownShareOneBulletAbstract()
    {
        // Mixed paths: an IEnumerable<string> auto-bullet (TokenValueHelpers.BulletList) and a
        // [Markdown] bullet list (OpenXmlMarkdownRenderer) must share the same bullet abstract.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Tags }}

            {{ Body }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MixedDoc>(template);

        var model = new MixedDoc
        {
            Tags = ["t1", "t2"],
            Body = "- m1\n- m2"
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 2);
    }

    [Test]
    public async Task LoopAndOuterListsShareOneBulletAbstract()
    {
        // The cloned ScopeTreeRunner inside a {% for %} loop body must inherit the parent runner's
        // WordNumberingState. The loop body uses the explicit `bullet_list` filter (auto-bullet
        // rendering skips loop scope), and combined with the outer auto-bullet list there should be
        // 3 NumberingInstances (2 from the loop iterations + 1 outer) all pointing at the same
        // shared bullet abstract.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for group in Groups %}

            {{ group.Items | bullet_list }}

            {% endfor %}

            {{ Outer }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<GroupsDoc>(template);

        var model = new GroupsDoc
        {
            Groups =
            [
                new() { Items = ["g1a", "g1b"] },
                new() { Items = ["g2a", "g2b"] }
            ],
            Outer = ["o1", "o2"]
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 3);
    }

    [Test]
    public async Task IfBranchAndOuterListsShareOneBulletAbstract()
    {
        // The inner ScopeTreeRunner spawned for {% if %} branch evaluation must inherit the parent
        // runner's WordNumberingState. One inside-branch list + one outside list = 2 bullet
        // NumberingInstances all pointing at the same bullet abstract.
        using var template = DocxTemplateBuilder.Build(
            """
            {% if ShowInner %}

            {{ Inner }}

            {% endif %}

            {{ Outer }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<IfDoc>(template);

        var model = new IfDoc
        {
            ShowInner = true,
            Inner = ["i1", "i2"],
            Outer = ["o1", "o2"]
        };

        await AssertOneSharedBulletAbstract(store, model, expectedInstanceCount: 2);
    }

    [Test]
    public async Task TwoMarkdownOrderedListsShareOneOrderedAbstract()
    {
        // The shared-state guarantee covers ordered lists too — WordNumberingState caches one
        // abstract per NumberFormatValues, so two markdown-rendered ordered lists must share.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ FirstBody }}

            {{ SecondBody }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TwoMarkdownDoc>(template);

        var model = new TwoMarkdownDoc
        {
            FirstBody = "1. a1\n2. a2",
            SecondBody = "1. b1\n2. b2"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var orderedDecimalAbstracts = numbering
            .Elements<AbstractNum>()
            .Where(a =>
                a.Elements<Level>().FirstOrDefault(_ => _.LevelIndex?.Value == 0) is { } lvl &&
                lvl.NumberingFormat?.Val?.Value == NumberFormatValues.Decimal)
            .ToList();
        var instancesPointingAtOrdered = numbering
            .Elements<NumberingInstance>()
            .Where(i =>
                i.AbstractNumId?.Val?.Value is { } refId &&
                orderedDecimalAbstracts.Any(a => a.AbstractNumberId?.Value == refId))
            .ToList();

        await Assert.That(orderedDecimalAbstracts.Count).IsEqualTo(1);
        await Assert.That(instancesPointingAtOrdered.Count).IsEqualTo(2);
    }

    static async Task AssertOneSharedBulletAbstract<TModel>(
        TemplateStore store,
        TModel model,
        int expectedInstanceCount)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var bulletAbstracts = numbering
            .Elements<AbstractNum>()
            .Where(a =>
                a.Elements<Level>().FirstOrDefault(_ => _.LevelIndex?.Value == 0) is { } lvl &&
                lvl.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
            .ToList();
        var instancesPointingAtBullet = numbering
            .Elements<NumberingInstance>()
            .Where(i =>
                i.AbstractNumId?.Val?.Value is { } refId &&
                bulletAbstracts.Any(a => a.AbstractNumberId?.Value == refId))
            .ToList();

        await Assert.That(bulletAbstracts.Count).IsEqualTo(1);
        await Assert.That(instancesPointingAtBullet.Count).IsEqualTo(expectedInstanceCount);
    }
}
