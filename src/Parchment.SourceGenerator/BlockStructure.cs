/// <summary>
/// SG-side mirror of the nesting rules in <c>Parchment.Tokens.ScopeTreeBuilder</c>. The builder
/// itself cannot be linked in: it pairs tags while building nodes over OpenXml paragraphs, which
/// the generator does not load. Same walk, same order, same wording, with the block tags of one
/// part as input.
/// </summary>
/// <remarks>
/// The builder's "Missing {% endfor %}." / "Missing {% endif %}." and trailing "Unexpected block
/// tag" throws have no counterpart: a body walk only returns at its closing tag, so nothing
/// reaches them. <c>StructuralLockstepTests</c> on each side pin the two to one case list.
/// </remarks>
static class BlockStructure
{
    public static string? FindError(IReadOnlyList<Token> tags)
    {
        var index = 0;
        return WalkBlock(tags, ref index, null);
    }

    static string? WalkBlock(IReadOnlyList<Token> tags, ref int index, TokenKind? closer)
    {
        while (index < tags.Count)
        {
            var tag = tags[index];
            if (closer != null &&
                IsCloserOrAlternate(tag.Kind, closer.Value))
            {
                return null;
            }

            string? error;
            switch (tag.Kind)
            {
                case TokenKind.ForOpen:
                    index++;
                    error = WalkBlock(tags, ref index, TokenKind.ForClose);
                    // Past the closing tag the body walk stopped at.
                    index++;
                    break;
                case TokenKind.IfOpen:
                    index++;
                    error = WalkIf(tags, ref index);
                    break;
                default:
                    return $"Unexpected '{tag.Source}' without a matching opening tag.";
            }

            if (error != null)
            {
                return error;
            }
        }

        if (closer == null)
        {
            return null;
        }

        if (closer == TokenKind.ForClose)
        {
            return "Missing closing tag 'endfor'.";
        }

        return "Missing closing tag 'endif'.";
    }

    static string? WalkIf(IReadOnlyList<Token> tags, ref int index)
    {
        var error = WalkBlock(tags, ref index, TokenKind.IfClose);
        if (error != null)
        {
            return error;
        }

        var seenElse = false;
        while (tags[index].Kind is TokenKind.ElsIf or TokenKind.Else)
        {
            var branch = tags[index];
            index++;

            // Checked before the branch body is walked, as the builder does.
            if (seenElse)
            {
                return $"Unexpected '{branch.Source}' after '{{% else %}}'.";
            }

            error = WalkBlock(tags, ref index, TokenKind.IfClose);
            if (error != null)
            {
                return error;
            }

            seenElse = branch.Kind == TokenKind.Else;
        }

        index++;
        return null;
    }

    static bool IsCloserOrAlternate(TokenKind kind, TokenKind closer) =>
        kind == closer ||
        (closer == TokenKind.IfClose && kind is TokenKind.ElsIf or TokenKind.Else);
}
