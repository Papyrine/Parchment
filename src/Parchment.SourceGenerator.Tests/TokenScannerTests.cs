public class TokenScannerTests
{
    [Test]
    public async Task SimpleSubstitution()
    {
        var tokens = TokenScanner.Scan(["Hello {{ customer.name }}!"]);
        await Assert.That(tokens.Count).IsEqualTo(1);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.Substitution);
        await Assert.That(tokens[0].References.Count).IsEqualTo(1);
    }

    [Test]
    public async Task ForLoop()
    {
        var tokens = TokenScanner.Scan([
            "{% for line in lines %}",
            "{{ line.description }}",
            "{% endfor %}"
        ]);
        await Assert.That(tokens.Count).IsEqualTo(3);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.ForOpen);
        await Assert.That(tokens[0].LoopVariable).IsEqualTo("line");
        await Assert.That(tokens[2].Kind).IsEqualTo(TokenKind.ForClose);
    }

    // A range's references are its bounds, so the token has to say that none is the loop source.
    [Test]
    public async Task ForLoop_Range()
    {
        var tokens = TokenScanner.Scan([
            "{% for i in (1..count) %}",
            "{% for line in lines %}"
        ]);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.ForOpen);
        await Assert.That(tokens[0].LoopVariable).IsEqualTo("i");
        await Assert.That(tokens[0].LoopSourceIsRange).IsTrue();
        await Assert.That(tokens[0].References.Count).IsEqualTo(1);
        await Assert.That(tokens[1].LoopSourceIsRange).IsFalse();
    }

    [Test]
    public async Task IfConditional()
    {
        var tokens = TokenScanner.Scan([
            "{% if customer.is_preferred %}",
            "Preferred!",
            "{% endif %}"
        ]);
        await Assert.That(tokens.Count).IsEqualTo(2);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.IfOpen);
        await Assert.That(tokens[1].Kind).IsEqualTo(TokenKind.IfClose);
    }

    [Test]
    public async Task UnknownBlockTag()
    {
        var tokens = TokenScanner.Scan(["{% foobar %}"]);
        await Assert.That(tokens.Count).IsEqualTo(1);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.UnknownBlock);
    }

    // The source has to stay References[0] whatever else the tag carries, since that is what the
    // generator resolves the loop variable from.
    [Test]
    public async Task ForLoopModifiers_AreKeptApartFromTheSource()
    {
        var tokens = TokenScanner.Scan(["{% for line in lines limit: max offset: skip reversed %}"]);
        await Assert.That(tokens[0].Kind).IsEqualTo(TokenKind.ForOpen);
        await Assert.That(tokens[0].References.Count).IsEqualTo(1);
        await Assert.That(tokens[0].References[0][0]).IsEqualTo("lines");
        await Assert.That(tokens[0].HasLoopModifiers).IsTrue();
        var modifiers = tokens[0].LoopModifierReferences!.Select(_ => _[0]).ToList();
        await Assert.That(modifiers).IsEquivalentTo(["max", "skip"]);
    }

    [Test]
    [Arguments("{% for line in lines %}", false)]
    [Arguments("{% for line in lines limit:1 %}", true)]
    [Arguments("{% for line in lines offset:1 %}", true)]
    [Arguments("{% for line in lines reversed %}", true)]
    public async Task ForLoopModifiers_AreDetected(string tag, bool expected)
    {
        var tokens = TokenScanner.Scan([tag]);
        await Assert.That(tokens[0].HasLoopModifiers).IsEqualTo(expected);
    }
}
