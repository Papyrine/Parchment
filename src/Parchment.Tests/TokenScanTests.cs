public class TokenScanTests
{
    [Test]
    public async Task EmptyText_NoSites()
    {
        var sites = TokenScan.Scan("");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task TextShorterThanMinimumToken_NoSites()
    {
        // Minimum token site is 4 chars (`{{}}` or `{%%}`).
        var sites = TokenScan.Scan("{{}");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task PlainText_NoSites()
    {
        var sites = TokenScan.Scan("just some normal text without tokens");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task SingleSubstitution()
    {
        var sites = TokenScan.Scan("{{ Customer.Name }}");
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(sites[0].Kind).IsEqualTo(TokenSiteKind.Substitution);
        await Assert.That(sites[0].Offset).IsEqualTo(0);
        await Assert.That(sites[0].Length).IsEqualTo(19);
    }

    [Test]
    public async Task SingleBlock()
    {
        var sites = TokenScan.Scan("{% if Customer.IsPreferred %}");
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(sites[0].Kind).IsEqualTo(TokenSiteKind.Block);
        await Assert.That(sites[0].Offset).IsEqualTo(0);
        await Assert.That(sites[0].Length).IsEqualTo(29);
    }

    [Test]
    public async Task SubstitutionInsideStaticText_OffsetReportedCorrectly()
    {
        const string text = "Hello {{ Name }}, welcome";
        var sites = TokenScan.Scan(text);
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(sites[0].Offset).IsEqualTo(6);
        await Assert.That(sites[0].Length).IsEqualTo(10);
        await Assert.That(text.Substring(sites[0].Offset, sites[0].Length)).IsEqualTo("{{ Name }}");
    }

    [Test]
    public async Task EmptySubstitutionBody()
    {
        // Mirrors regex behavior: `[^{}]*?` allows zero chars, so `{{}}` is a valid match.
        var sites = TokenScan.Scan("{{}}");
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(sites[0].Kind).IsEqualTo(TokenSiteKind.Substitution);
        await Assert.That(sites[0].Length).IsEqualTo(4);
    }

    [Test]
    public async Task EmptyBlockBody()
    {
        var sites = TokenScan.Scan("{%%}");
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(sites[0].Kind).IsEqualTo(TokenSiteKind.Block);
        await Assert.That(sites[0].Length).IsEqualTo(4);
    }

    [Test]
    public async Task AdjacentSubstitutions()
    {
        var sites = TokenScan.Scan("{{a}}{{b}}");
        await Assert.That(sites.Count).IsEqualTo(2);
        await Assert.That(sites[0].Offset).IsEqualTo(0);
        await Assert.That(sites[0].Length).IsEqualTo(5);
        await Assert.That(sites[1].Offset).IsEqualTo(5);
        await Assert.That(sites[1].Length).IsEqualTo(5);
    }

    [Test]
    public async Task MixedSubstitutionAndBlockOnSameLine()
    {
        const string text = "{% if x %} value is {{ x }} {% endif %}";
        var sites = TokenScan.Scan(text);
        await Assert.That(sites.Count).IsEqualTo(3);
        await Assert.That(sites[0].Kind).IsEqualTo(TokenSiteKind.Block);
        await Assert.That(sites[1].Kind).IsEqualTo(TokenSiteKind.Substitution);
        await Assert.That(sites[2].Kind).IsEqualTo(TokenSiteKind.Block);
        await Assert.That(text.Substring(sites[0].Offset, sites[0].Length)).IsEqualTo("{% if x %}");
        await Assert.That(text.Substring(sites[1].Offset, sites[1].Length)).IsEqualTo("{{ x }}");
        await Assert.That(text.Substring(sites[2].Offset, sites[2].Length)).IsEqualTo("{% endif %}");
    }

    [Test]
    public async Task IncompleteSubstitution_NoMatch()
    {
        var sites = TokenScan.Scan("{{ Name");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task IncompleteBlock_NoMatch()
    {
        var sites = TokenScan.Scan("{% if x");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task SubstitutionBodyContainingBrace_NoMatch()
    {
        // Body cannot contain `{` — mirrors regex char class `[^{}]`.
        var sites = TokenScan.Scan("{{ a{b }}");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task SubstitutionWithSingleBrace_NoMatch()
    {
        // A lone `}` in the body terminates the body without a paired `}}` — no match.
        var sites = TokenScan.Scan("{{ a}b }}");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task BlockBodyContainingBrace_NoMatch()
    {
        // Body cannot contain `{` — mirrors regex char class `[^{%]`.
        var sites = TokenScan.Scan("{% if {x %}");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task BlockBodyContainingPercent_NoMatch()
    {
        // A lone `%` in the body terminates the body without a paired `%}` — no match.
        var sites = TokenScan.Scan("{% if x%y %}");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task ScannerSkipsPastFalseStartAndKeepsScanning()
    {
        // The leading `{{` doesn't form a valid substitution (body has `{`),
        // but the scanner must keep going and find the real one later.
        const string text = "{{ a{b }} ok {{ valid }}";
        var sites = TokenScan.Scan(text);
        await Assert.That(sites.Count).IsEqualTo(1);
        await Assert.That(text.Substring(sites[0].Offset, sites[0].Length)).IsEqualTo("{{ valid }}");
    }

    [Test]
    public async Task LoneBraceCharacters_Ignored()
    {
        var sites = TokenScan.Scan("a { b } c");
        await Assert.That(sites).IsEmpty();
    }

    [Test]
    public async Task HasContentOutsideSites_EmptyText_False()
    {
        var sites = TokenScan.Scan("");
        await Assert.That(TokenScan.HasContentOutsideSites("", sites)).IsFalse();
    }

    [Test]
    public async Task HasContentOutsideSites_OnlyWhitespace_False()
    {
        var sites = TokenScan.Scan("   ");
        await Assert.That(TokenScan.HasContentOutsideSites("   ", sites)).IsFalse();
    }

    [Test]
    public async Task HasContentOutsideSites_PlainTextNoSites_True()
    {
        var sites = TokenScan.Scan("hello");
        await Assert.That(TokenScan.HasContentOutsideSites("hello", sites)).IsTrue();
    }

    [Test]
    public async Task HasContentOutsideSites_OnlyTokenWithWhitespacePadding_False()
    {
        const string text = "  {{ x }}  ";
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.HasContentOutsideSites(text, sites)).IsFalse();
    }

    [Test]
    public async Task HasContentOutsideSites_NonWhitespaceBeforeToken_True()
    {
        const string text = "Hello {{ x }}";
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.HasContentOutsideSites(text, sites)).IsTrue();
    }

    [Test]
    public async Task HasContentOutsideSites_NonWhitespaceAfterToken_True()
    {
        const string text = "{{ x }} world";
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.HasContentOutsideSites(text, sites)).IsTrue();
    }

    [Test]
    public async Task HasContentOutsideSites_NonWhitespaceBetweenTokens_True()
    {
        const string text = "{% if x %} value {% endif %}";
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.HasContentOutsideSites(text, sites)).IsTrue();
    }

    [Test]
    public async Task HasContentOutsideSites_OnlyWhitespaceBetweenTokens_False()
    {
        const string text = "{{ a }}   {{ b }}";
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.HasContentOutsideSites(text, sites)).IsFalse();
    }

    [Test]
    [Arguments("{{ x }}", true)]
    [Arguments("  {{ x }}\t", true)]
    [Arguments("\u00A0{{ x }}\n", true)]
    [Arguments("a {{ x }}", false)]
    [Arguments("{{ x }} a", false)]
    [Arguments("{{ x }} {{ y }}", false)]
    public async Task IsSolo(string text, bool expected)
    {
        var sites = TokenScan.Scan(text);
        await Assert.That(TokenScan.IsSolo(text, sites.Count, sites[0].Offset, sites[0].Length)).IsEqualTo(expected);
    }
}
