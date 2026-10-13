// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A token names a model member without regard to case. Registration accepts
/// <c>{{ customer.name }}</c> for <c>Customer.Name</c>, so the render has to bind it as well: in
/// both flows, at every depth, and through every map that is keyed by member path.
/// </summary>
public partial class MemberCaseTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public required string Name { get; init; }
        public bool Flag { get; init; }
        public required string Code;
        public static string Shared => "shared";
        public required Customer Customer { get; init; }
        public required List<Line> Lines { get; init; }
        public required string[] Tags { get; init; }

        [Html]
        public required string Body { get; init; }

        [Markdown]
        public required string Notes { get; init; }

        [ExcelsiorTable]
        public required List<Line> Rows { get; init; }
    }

    public class Customer
    {
        public required string Name { get; init; }
        public int Limit { get; init; }
        public required Address Address { get; init; }
        public required List<Line> Orders { get; init; }
        public required List<string> Skills { get; init; }

        [Html]
        public required string Bio { get; init; }
    }

    public class Address
    {
        public required string City { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
        public required int Quantity { get; init; }
    }

    [ParchmentBindable]
    public partial class Form
    {
        public required string Title { get; init; }

        [EditableField]
        public required string Reference { get; set; }

        [Html]
        [EditableField]
        public required string Summary { get; set; }

        public required Contact Contact { get; init; }

        [EditableField]
        public required List<Bid> Bids { get; set; }
    }

    public class Contact
    {
        [EditableField]
        public required string Email { get; set; }
    }

    public class Bid
    {
        public string Label { get; init; } = "bid";

        [EditableField]
        public required string Year { get; set; }

        [EditableField]
        public required decimal Amount { get; set; }
    }

    static List<Line> BuildLines() =>
    [
        new()
        {
            Name = "one",
            Quantity = 1
        },
        new()
        {
            Name = "two",
            Quantity = 2
        }
    ];

    static Doc BuildModel() =>
        new()
        {
            Name = "Alice",
            Flag = true,
            Code = "C1",
            Customer = new()
            {
                Name = "Acme",
                Limit = 1,
                Address = new()
                {
                    City = "Perth"
                },
                Orders = BuildLines(),
                Skills = ["weld", "paint"],
                Bio = "<p>a <b>bold</b> bio</p>"
            },
            Lines = BuildLines(),
            Tags = ["red", "green"],
            Body = "<p>first</p><p>second <i>italic</i></p>",
            Notes = "a **strong** note",
            Rows = BuildLines()
        };

    static Form BuildForm() =>
        new()
        {
            Title = "Tender",
            Reference = "REF-7",
            Summary = "<p>a <b>rich</b> summary</p>",
            Contact = new()
            {
                Email = "bids@example.com"
            },
            Bids =
            [
                new()
                {
                    Year = "2026",
                    Amount = 10m
                },
                new()
                {
                    Year = "2027",
                    Amount = 20m
                }
            ]
        };

    [Test]
    [Arguments("{{ name }}", "Alice")]
    [Arguments("{{ NAME }}", "Alice")]
    [Arguments("{{ nAmE | upcase }}", "ALICE")]
    [Arguments("{{ code }}", "C1")]
    [Arguments("{{ shared }}", "shared")]
    [Arguments("{{ customer.Name }}", "Acme")]
    [Arguments("{{ Customer.name }}", "Acme")]
    [Arguments("{{ CUSTOMER.NAME }}", "Acme")]
    [Arguments("{{ customer['name'] }}", "Acme")]
    [Arguments("{{ Customer.address.city }}", "Perth")]
    [Arguments("{{ lines.size }}", "2")]
    [Arguments("{{ lines.first.name }}", "one")]
    [Arguments("{{ customer.orders.last.NAME }}", "two")]
    [Arguments("{{ Lines | map: 'name' | join: ' ' }}", "one two")]
    public async Task Substitution(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    [Test]
    [Arguments("flag")]
    [Arguments("FLAG")]
    [Arguments("customer.name == 'Acme'")]
    [Arguments("Customer.address.CITY != 'Sydney'")]
    [Arguments("lines.size > 1")]
    [Arguments("tags contains 'red'")]
    public async Task Condition(string condition)
    {
        var docx = await RenderDocx(
            $$"""
              {% if {{condition}} %}

              yes

              {% else %}

              no

              {% endif %}
              """);
        await Assert.That(docx).IsEqualTo("yes");

        var markdown = await RenderMarkdown($"{{% if {condition} %}}yes{{% else %}}no{{% endif %}}");
        await Assert.That(markdown).IsEqualTo("yes");
    }

    [Test]
    public async Task ElsifCondition()
    {
        var docx = await RenderDocx(
            """
            {% if customer.name == 'Other' %}

            first

            {% elsif flag %}

            second

            {% endif %}
            """);
        await Assert.That(docx).IsEqualTo("second");

        var markdown = await RenderMarkdown("{% if customer.name == 'Other' %}first{% elsif flag %}second{% endif %}");
        await Assert.That(markdown).IsEqualTo("second");
    }

    [Test]
    [Arguments("lines")]
    [Arguments("LINES")]
    [Arguments("customer.orders")]
    [Arguments("Customer.ORDERS")]
    public async Task LoopSourceAndItemMember(string source)
    {
        var docx = await RenderDocx(
            $$$"""
               {% for line in {{{source}}} %}

               {{ line.name }}={{ line.QUANTITY }};

               {% endfor %}
               """);
        await Assert.That(docx).IsEqualTo("one=1;two=2;");

        var markdown = await RenderMarkdown($"{{% for line in {source} %}}{{{{ line.name }}}}={{{{ line.QUANTITY }}}};{{% endfor %}}");
        await Assert.That(markdown).IsEqualTo("one=1;two=2;");
    }

    [Test]
    public async Task LoopModifier()
    {
        var docx = await RenderDocx(
            """
            {% for line in lines limit: customer.limit %}

            {{ line.name }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("one;");

        var markdown = await RenderMarkdown("{% for line in lines limit: customer.limit %}{{ line.name }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("one;");
    }

    [Test]
    public async Task ConditionalInsideALoop()
    {
        var docx = await RenderDocx(
            """
            {% for line in lines %}

            {% if flag %}

            {{ line.name }} of {{ customer.name }};

            {% endif %}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("one of Acme;two of Acme;");

        var markdown = await RenderMarkdown("{% for line in lines %}{% if flag %}{{ line.name }} of {{ customer.name }};{% endif %}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("one of Acme;two of Acme;");
    }

    // A loop variable is still found by its exact spelling, so a token in another case is the
    // root member and not the item.
    [Test]
    public async Task LoopVariableKeepsItsExactSpelling()
    {
        var docx = await RenderDocx(
            """
            {% for name in tags %}

            {{ name }}/{{ Name }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("red/Alice;green/Alice;");

        var markdown = await RenderMarkdown("{% for name in tags %}{{ name }}/{{ Name }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("red/Alice;green/Alice;");
    }

    [Test]
    public async Task HeaderAndFooter()
    {
        using var template = BuildWithHeaderAndFooter("{{ name }}", "{{ customer.name }}", "{{ Customer.address.CITY }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        await AssertParts(store);

        using var styleSource = BuildWithHeaderAndFooter("replaced", "{{ customer.name }}", "{{ Customer.address.CITY }}");
        store = new();
        store.RegisterMarkdownTemplate<Doc>("{{ name }}", styleSource);
        await AssertParts(store);

        static async Task AssertParts(TemplateStore store)
        {
            using var stream = await Render(store, BuildModel());
            using var doc = WordprocessingDocument.Open(stream, false);
            var mainPart = doc.MainDocumentPart!;
            await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Alice");
            await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("Acme");
            await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo("Perth");
        }
    }

    // What the member's attribute or type makes of the token is decided from maps keyed by the
    // member's path, so the document has to come out the same as for the path as declared.
    [Test]
    [Arguments("{{ body }}", "{{ Body }}", "firstsecond italic")]
    [Arguments("before {{ BODY }} after", "before {{ Body }} after", "before firstsecond italic after")]
    [Arguments("{{ notes }}", "{{ Notes }}", "a strong note")]
    [Arguments("{{ customer.bio }}", "{{ Customer.Bio }}", "a bold bio")]
    [Arguments("{{ Customer.BIO }} and {{ name }}", "{{ Customer.Bio }} and {{ Name }}", "a bold bio and Alice")]
    [Arguments("{{ tags }}", "{{ Tags }}", "redgreen")]
    [Arguments("Tags: {{ tags }}", "Tags: {{ Tags }}", "Tags: redgreen")]
    [Arguments("{{ customer.skills }}", "{{ Customer.Skills }}", "weldpaint")]
    [Arguments("{{ rows }}", "{{ Rows }}", "NameQuantityone1two2")]
    public async Task AttributedMemberInDocx(string token, string declared, string text)
    {
        var expected = await RenderDocxBody(declared, BuildModel());
        var actual = await RenderDocxBody(token, BuildModel());

        await Assert.That(expected.InnerText).IsEqualTo(text);
        await Assert.That(actual.OuterXml).IsEqualTo(expected.OuterXml);
    }

    [Test]
    [Arguments("{{ body }}", "{{ Body }}", "firstsecond italic")]
    [Arguments("{{ notes }}", "{{ Notes }}", "a strong note")]
    [Arguments("{{ customer.bio }}", "{{ Customer.Bio }}", "a bold bio")]
    [Arguments("{{ tags }}", "{{ Tags }}", "redgreen")]
    [Arguments("{{ customer.skills | bullet_list }}", "{{ Customer.Skills | bullet_list }}", "weldpaint")]
    [Arguments("{{ rows }}", "{{ Rows }}", "NameQuantityone1two2")]
    public async Task AttributedMemberInMarkdown(string token, string declared, string text)
    {
        var expected = await RenderMarkdownBody(declared);
        var actual = await RenderMarkdownBody(token);

        await Assert.That(expected.InnerText).IsEqualTo(text);
        await Assert.That(actual.OuterXml).IsEqualTo(expected.OuterXml);
    }

    // The control's tag is the member's path as declared, whatever the token spelled, so the
    // document extracts the same.
    [Test]
    [Arguments("{{ title }}: {{ reference }}", "{{ Title }}: {{ Reference }}", "Tender: REF-7")]
    [Arguments("{{ contact.email }}", "{{ Contact.Email }}", "bids@example.com")]
    [Arguments("{{ summary }}", "{{ Summary }}", "a rich summary")]
    public async Task EditableField(string token, string declared, string text)
    {
        var expected = await RenderDocxBody(declared, BuildForm());
        var actual = await RenderDocxBody(token, BuildForm());

        await Assert.That(expected.Descendants<SdtElement>().Count()).IsEqualTo(1);
        await Assert.That(expected.InnerText).IsEqualTo(text);
        await Assert.That(actual.OuterXml).IsEqualTo(expected.OuterXml);
    }

    // The loop source decides both that the loop is a repeating section and what its rows are. Only
    // the first of those went through a map, so a source in another case rendered the section with
    // no rows, and extraction then replaced the collection with nothing.
    [Test]
    public async Task EditableCollection()
    {
        const string declared =
            """
            {% for bid in Bids %}

            {{ bid.Label }} {{ bid.Year }}: {{ bid.Amount }}

            {% endfor %}
            """;
        const string template =
            """
            {% for bid in bids %}

            {{ bid.label }} {{ bid.year }}: {{ bid.AMOUNT }}

            {% endfor %}
            """;

        var expected = await RenderDocxBody(declared, BuildForm());
        var actual = await RenderDocxBody(template, BuildForm());

        await Assert.That(expected.InnerText).IsEqualTo("bid 2026: 10bid 2027: 20");
        await Assert.That(actual.OuterXml).IsEqualTo(expected.OuterXml);

        using var docx = DocxTemplateBuilder.Build(template);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Form>(docx);
        using var stream = await Render(store, BuildForm());

        var extracted = new Form
        {
            Title = "x",
            Reference = "x",
            Summary = "x",
            Contact = new()
            {
                Email = "x"
            },
            Bids = []
        };
        ParchmentExtractor.Extract<Form>(stream).ApplyTo(extracted);

        await Assert.That(extracted.Bids.Select(_ => _.Year)).IsEquivalentTo(["2026", "2027"]);
        await Assert.That(extracted.Bids.Select(_ => _.Amount)).IsEquivalentTo([10m, 20m]);
    }

    // The rules a member's attribute puts on its token are found through the same maps, so they
    // hold for a token in any case as well.
    [Test]
    [Arguments("Rows: {{ rows }}")]
    [Arguments("{{ ROWS | reverse }}")]
    public async Task ExcelsiorTokenRulesStillApply(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("[ExcelsiorTable]");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(content, styleSource));
        await Assert.That(markdown.Message).Contains("[ExcelsiorTable]");
    }

    [Test]
    public async Task FormatAndEditableTokenRulesStillApply()
    {
        using var filtered = DocxTemplateBuilder.Build("{{ body | upcase }}");
        var format = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(filtered));
        await Assert.That(format.Message).Contains("{{ body | upcase }}");

        using var repeated = DocxTemplateBuilder.Build("{{ Reference }} and {{ reference }}");
        var editable = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Form>(repeated));
        await Assert.That(editable.Message).Contains("more than once");
    }

    static async Task<string> RenderDocx(string content) =>
        (await RenderDocxBody(content, BuildModel())).InnerText;

    static async Task<string> RenderMarkdown(string markdown) =>
        (await RenderMarkdownBody(markdown)).InnerText;

    static async Task<Body> RenderDocxBody<T>(string content, T model)
        where T : notnull
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);
        return await RenderBody(store, model);
    }

    static async Task<Body> RenderMarkdownBody(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await RenderBody(store, BuildModel());
    }

    static async Task<Body> RenderBody<T>(TemplateStore store, T model)
        where T : notnull
    {
        using var stream = await Render(store, model);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static async Task<MemoryStream> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildWithHeaderAndFooter(string body, string header, string footer)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(Paragraph(header));

            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
            footerPart.Footer = new(Paragraph(footer));

            mainPart.Document.Body!.Append(
                Paragraph(body),
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter"
                    },
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
        }

        stream.Position = 0;
        return stream;

        static Paragraph Paragraph(string text) =>
            new(new Run(new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            }));
    }
}
