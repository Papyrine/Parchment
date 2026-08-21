public class XmlCharsTests
{
    [Test]
    public async Task Empty_ReturnsEmpty()
    {
        var result = XmlChars.Strip("").ToString();
        await Assert.That(result).IsEqualTo("");
    }

    [Test]
    public async Task PlainAscii_ReturnsSameSpan()
    {
        var input = "Plain ASCII text 123";
        var inputSpan = input.AsSpan();
        // Fast path must avoid copying: returns the original span unchanged (same memory).
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task WhitespaceControls_PreservedByFastPath()
    {
        var input = "tab\there\nline\rreturn";
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task BmpUnicode_ReturnsSameSpan()
    {
        // Unicode in [0x20, 0xD7FF] and [0xE000, 0xFFFD] is XML-valid and must hit the fast path.
        var input = "café 日本語 résumé";
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task NullChar_Stripped()
    {
        var result = XmlChars.Strip("a\0b").ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task LowControlsExceptWhitespace_AllStripped()
    {
        // 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F are XML 1.0 forbidden (except tab=0x09, LF=0x0A, CR=0x0D).
        var builder = new StringBuilder();
        builder.Append('a');
        for (var c = 0x00; c <= 0x08; c++)
        {
            builder.Append((char)c);
        }
        builder.Append('b');
        builder.Append((char)0x0B);
        builder.Append((char)0x0C);
        builder.Append('c');
        for (var c = 0x0E; c <= 0x1F; c++)
        {
            builder.Append((char)c);
        }
        builder.Append('d');

        var result = XmlChars.Strip(builder.ToString()).ToString();
        await Assert.That(result).IsEqualTo("abcd");
    }

    [Test]
    public async Task TabLfCr_Preserved()
    {
        var input = "x\ty\nz\rw";
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo(input);
    }

    [Test]
    public async Task NonCharacterFFFE_Stripped()
    {
        var result = XmlChars.Strip("a￾b").ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task NonCharacterFFFF_Stripped()
    {
        var result = XmlChars.Strip("a￿b").ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task FFFD_Preserved()
    {
        // 0xFFFD (replacement character) is the upper edge of the second valid BMP range.
        var input = "a�b";
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo(input);
    }

    [Test]
    public async Task ValidSurrogatePair_Preserved()
    {
        // U+1F600 (😀) encodes as high D83D + low DE00 — a valid pair.
        // Surrogates are in the inspection set, so the fast-path's IndexOfAny matches and
        // hands off to slow path. Slow path must walk the pair, find it valid, and return
        // the original span unchanged (no allocation) — same-memory assertion guards
        // against a regression where the builder gets eagerly allocated on slow-path entry.
        var input = "smile 😀 here";
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task LoneHighSurrogate_Stripped()
    {
        var input = "a\uD83Db"; // high surrogate without a low follower
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task LoneLowSurrogate_Stripped()
    {
        var input = "a\uDE00b"; // low surrogate without a high predecessor
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task HighSurrogateAtEnd_Stripped()
    {
        var input = "abc\uD83D";
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo("abc");
    }

    [Test]
    public async Task HighSurrogateFollowedByNonLow_BothEvaluatedSeparately()
    {
        // High surrogate followed by a regular BMP char — high is lone (stripped),
        // the BMP char is valid (kept).
        var input = "\uD83Dx";
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo("x");
    }

    [Test]
    public async Task TwoConsecutiveHighSurrogates_BothStripped()
    {
        // Two highs in a row: the first is lone (stripped), then the second is also lone (stripped).
        var input = "a\uD83D\uD83Db";
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo("ab");
    }

    [Test]
    public async Task MixedValidAndInvalid_OnlyInvalidStripped()
    {
        // tab + null + 'A' + lone high surrogate + valid pair + 0xFFFE + 'B'
        var input = "\t\0A\uD83D😀￾B";
        var result = XmlChars.Strip(input).ToString();
        // Expected: tab + 'A' + valid pair (smile) + 'B'
        await Assert.That(result).IsEqualTo("\tA😀B");
    }

    [Test]
    public async Task LongCleanString_FastPathReturnsSameSpan()
    {
        var input = new string('a', 10_000);
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task SingleInvalidCharInLongString_RestPreserved()
    {
        var prefix = new string('a', 5_000);
        var suffix = new string('b', 5_000);
        var input = prefix + "\0" + suffix;
        var result = XmlChars.Strip(input).ToString();
        await Assert.That(result).IsEqualTo(prefix + suffix);
    }

    [Test]
    public async Task SingleChar_Tab_Preserved()
    {
        var result = XmlChars.Strip("\t").ToString();
        await Assert.That(result).IsEqualTo("\t");
    }

    [Test]
    public async Task SingleChar_Null_Stripped()
    {
        var result = XmlChars.Strip("\0").ToString();
        await Assert.That(result).IsEqualTo("");
    }

    [Test]
    public async Task BoundaryD7FF_Preserved()
    {
        // Last char before the surrogate range — must hit fast path (not in inspection set).
        var input = "a퟿b";
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }

    [Test]
    public async Task BoundaryE000_Preserved()
    {
        // First char after the surrogate range.
        var input = "ab";
        var inputSpan = input.AsSpan();
        var sameMemory = XmlChars.Strip(inputSpan) == inputSpan;
        await Assert.That(sameMemory).IsTrue();
    }
}
