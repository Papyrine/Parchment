/// <summary>
/// Builds a scope tree from a flat list of paragraph classifications. Opening block tags (for/if)
/// nest their following paragraphs as body nodes up to their matching closing tag.
/// </summary>
static class ScopeTreeBuilder
{
    public static IReadOnlyList<RangeNode> Build(
        IReadOnlyList<ParagraphClassification> classifications,
        string templateName,
        string partUri)
    {
        var queue = new Queue<ParagraphClassification>(classifications);
        var nodes = BuildBlock(queue, null, templateName, partUri);
        if (queue.Count == 0)
        {
            return nodes;
        }

        var extra = queue.Peek();
        throw new ParchmentRegistrationException(
            templateName,
            $"Unexpected block tag '{extra.Block?.Source}' without a matching opening.",
            partUri,
            extra.Block?.Source);
    }

    static List<RangeNode> BuildBlock(
        Queue<ParagraphClassification> queue,
        BlockTagKind? closer,
        string templateName,
        string partUri)
    {
        var result = new List<RangeNode>();
        while (queue.Count > 0)
        {
            var next = queue.Peek();
            if (next is { Kind: ParagraphKind.Block, Block: not null })
            {
                var kind = next.Block.Kind;
                if (closer != null && IsCloserOrAlternate(kind, closer.Value))
                {
                    return result;
                }

                switch (kind)
                {
                    case BlockTagKind.For:
                        queue.Dequeue();
                        result.Add(BuildFor(next, queue, templateName, partUri));
                        break;
                    case BlockTagKind.If:
                        queue.Dequeue();
                        result.Add(BuildIf(next, queue, templateName, partUri));
                        break;
                    case BlockTagKind.EndFor:
                    case BlockTagKind.EndIf:
                    case BlockTagKind.ElsIf:
                    case BlockTagKind.Else:
                        throw new ParchmentRegistrationException(
                            templateName,
                            $"Unexpected '{next.Block.Source}' without a matching opening tag.",
                            partUri,
                            next.Block.Source);
                }

                continue;
            }

            queue.Dequeue();
            if (next.Kind == ParagraphKind.Substitution)
            {
                result.Add(new SubstitutionNode(next.AnchorName, next.Substitutions));
            }
            else
            {
                result.Add(new StaticNode(next.AnchorName));
            }
        }

        if (closer == null)
        {
            return result;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"Missing closing tag '{closer.Value.ToString().ToLowerInvariant()}'.",
            partUri);
    }

    static LoopNode BuildFor(
        ParagraphClassification opening,
        Queue<ParagraphClassification> queue,
        string templateName,
        string partUri)
    {
        var body = BuildBlock(queue, BlockTagKind.EndFor, templateName, partUri);
        if (queue.Count == 0 ||
            queue.Peek().Block?.Kind != BlockTagKind.EndFor)
        {
            throw new ParchmentRegistrationException(
                templateName,
                "Missing {% endfor %}.",
                partUri,
                opening.Block?.Source);
        }

        var closing = queue.Dequeue();
        var loopSource = opening.Block!.LoopSource!;
        // Resolve the source's first dotted reference once, here at build time, so ProcessLoopAsync
        // never has to run IdentifierVisitor.Collect (which allocates a visitor + path lists) per loop.
        var sourceRefs = IdentifierVisitor.Collect(loopSource);
        var sourceReference = sourceRefs.Count > 0 ? sourceRefs[0].Dotted : null;
        return new(
            opening.AnchorName,
            closing.AnchorName,
            opening.Block!.LoopVariable!,
            loopSource,
            sourceReference,
            body);
    }

    static IfNode BuildIf(
        ParagraphClassification opening,
        Queue<ParagraphClassification> queue,
        string templateName,
        string partUri)
    {
        var branches = new List<IfBranch>();
        var elseBody = new List<RangeNode>();
        string? elseAnchorName = null;

        // First branch is the if itself
        var firstBody = BuildBlock(queue, BlockTagKind.EndIf, templateName, partUri);
        branches.Add(new(opening.Block!.Condition!, opening.AnchorName, firstBody));

        // Collect elsif / else branches until endif
        while (queue.TryPeek(out var peek) &&
               peek.Block?.Kind is BlockTagKind.ElsIf or BlockTagKind.Else)
        {
            var branchOpening = queue.Dequeue();

            // Branch tags after {% else %} break the positional branch ranges the runner keeps —
            // and are malformed liquid anyway (the markdown flow's full-template Fluid parse
            // rejects them too).
            if (elseAnchorName != null)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"Unexpected '{branchOpening.Block!.Source}' after '{{% else %}}'.",
                    partUri,
                    branchOpening.Block!.Source);
            }

            var branchBody = BuildBlock(queue, BlockTagKind.EndIf, templateName, partUri);
            if (branchOpening.Block!.Kind == BlockTagKind.ElsIf)
            {
                branches.Add(new(branchOpening.Block!.Condition!, branchOpening.AnchorName, branchBody));
            }
            else
            {
                elseAnchorName = branchOpening.AnchorName;
                elseBody.AddRange(branchBody);
            }
        }

        if (queue.Count == 0 ||
            queue.Peek().Block?.Kind != BlockTagKind.EndIf)
        {
            throw new ParchmentRegistrationException(
                templateName,
                "Missing {% endif %}.",
                partUri,
                opening.Block?.Source);
        }

        var closing = queue.Dequeue();
        return new(
            opening.AnchorName,
            closing.AnchorName,
            branches,
            elseAnchorName,
            elseBody);
    }

    static bool IsCloserOrAlternate(BlockTagKind kind, BlockTagKind closer) =>
        kind == closer ||
        (closer == BlockTagKind.EndIf && kind is BlockTagKind.ElsIf or BlockTagKind.Else);
}
