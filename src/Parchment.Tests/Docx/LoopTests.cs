public class LoopTests
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
        store.RegisterDocxTemplate<Invoice>("loop", template);

        using var stream = new MemoryStream();
        await store.Render("loop", SampleData.Invoice(), stream);
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
        store.RegisterDocxTemplate<NestedModel>("nested-loop", template);
        using var stream = new MemoryStream();
        await store.Render(
            "nested-loop",
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

    public class NestedModel
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
        store.RegisterDocxTemplate<Invoice>("empty-loop", template);

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
        await store.Render("empty-loop", emptied, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class TripleNestModel
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
        store.RegisterDocxTemplate<TripleNestModel>("triple-nest", template);

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
        await store.Render("triple-nest", model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class LeakModel
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
        store.RegisterDocxTemplate<LeakModel>("loop-leak", template);

        using var stream = new MemoryStream();
        await store.Render(
            "loop-leak",
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

    public class NestedShadowModel
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
        store.RegisterDocxTemplate<NestedShadowModel>("loop-shadow", template);

        using var stream = new MemoryStream();
        await store.Render(
            "loop-shadow",
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
}
