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
        var textBoxHosts = TextBoxHosts(classifications);
        var nodes = BuildBlock(queue, null, textBoxHosts, templateName, partUri);
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
        HashSet<Paragraph> textBoxHosts,
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
                        result.Add(BuildFor(next, queue, textBoxHosts, templateName, partUri));
                        break;
                    case BlockTagKind.If:
                        queue.Dequeue();
                        result.Add(BuildIf(next, queue, textBoxHosts, templateName, partUri));
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
        HashSet<Paragraph> textBoxHosts,
        string templateName,
        string partUri)
    {
        var body = BuildBlock(queue, BlockTagKind.EndFor, textBoxHosts, templateName, partUri);
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
        RequireSameContainer([opening, closing], templateName, partUri);
        var loopSource = opening.Block!.LoopSource!;
        // Resolve the source's first dotted reference once, here at build time, so ProcessLoopAsync
        // never has to run IdentifierVisitor.Collect (which allocates a visitor + path lists) per loop.
        var sourceRefs = IdentifierVisitor.Collect(loopSource);
        // A range's references are its bounds, not what is iterated, so it has no source reference.
        var sourceReference = loopSource is MemberExpression && sourceRefs.Count > 0 ? sourceRefs[0].Dotted : null;
        return new(
            opening.AnchorName,
            closing.AnchorName,
            opening.Block!.LoopVariable!,
            loopSource,
            sourceReference,
            opening.Block.LoopLimit,
            opening.Block.LoopOffset,
            opening.Block.LoopReversed,
            HoldsTextBoxHost(opening, closing, textBoxHosts),
            body);
    }

    static IfNode BuildIf(
        ParagraphClassification opening,
        Queue<ParagraphClassification> queue,
        HashSet<Paragraph> textBoxHosts,
        string templateName,
        string partUri)
    {
        var branches = new List<IfBranch>();
        var tags = new List<ParagraphClassification>
        {
            opening
        };
        var elseBody = new List<RangeNode>();
        string? elseAnchorName = null;

        // First branch is the if itself
        var firstBody = BuildBlock(queue, BlockTagKind.EndIf, textBoxHosts, templateName, partUri);
        branches.Add(new(opening.Block!.Condition!, opening.AnchorName, firstBody));

        // Collect elsif / else branches until endif
        while (queue.TryPeek(out var peek) &&
               peek.Block?.Kind is BlockTagKind.ElsIf or BlockTagKind.Else)
        {
            var branchOpening = queue.Dequeue();
            tags.Add(branchOpening);

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

            var branchBody = BuildBlock(queue, BlockTagKind.EndIf, textBoxHosts, templateName, partUri);
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
        tags.Add(closing);
        RequireSameContainer(tags, templateName, partUri);
        return new(
            opening.AnchorName,
            closing.AnchorName,
            branches,
            elseAnchorName,
            elseBody,
            HoldsTextBoxHost(opening, closing, textBoxHosts));
    }

    // Every paragraph with an anchored paragraph somewhere below it, which is to say every
    // paragraph anchoring a text box that holds a token. Empty for a part without one, and then
    // no block is searched.
    static HashSet<Paragraph> TextBoxHosts(IReadOnlyList<ParagraphClassification> classifications)
    {
        var hosts = new HashSet<Paragraph>();
        foreach (var classification in classifications)
        {
            foreach (var ancestor in classification.Paragraph.Ancestors<Paragraph>())
            {
                hosts.Add(ancestor);
            }
        }

        return hosts;
    }

    // Whether one of the block's elements, the siblings between its tags that the runner clones
    // or keeps, is such a paragraph. The runner indexes a paragraph by its direct children and
    // every other element by all the paragraphs below it (ScopeTreeRunner.CollectAnchors), so
    // only a paragraph can hide an anchor from it. Called once the tags are known to be siblings.
    static bool HoldsTextBoxHost(ParagraphClassification opening, ParagraphClassification closing, HashSet<Paragraph> textBoxHosts)
    {
        if (textBoxHosts.Count == 0)
        {
            return false;
        }

        var cursor = opening.Paragraph.NextSibling();
        while (cursor != null &&
               cursor != closing.Paragraph)
        {
            if (cursor is Paragraph paragraph &&
                textBoxHosts.Contains(paragraph))
            {
                return true;
            }

            cursor = cursor.NextSibling();
        }

        return false;
    }

    // Tags are paired by document order, but the runner reaches a block's content by walking
    // siblings from one tag paragraph to the next (ScopeTreeRunner.CaptureBetween). A tag in a
    // different parent is never reached by that walk, so the block would swallow everything after
    // its open tag or leave its body unrendered. Lockstep with PARCH028 in the source generator.
    static void RequireSameContainer(
        List<ParagraphClassification> tags,
        string templateName,
        string partUri)
    {
        var opening = tags[0];
        foreach (var tag in tags)
        {
            if (ReferenceEquals(tag.Paragraph.Parent, opening.Paragraph.Parent))
            {
                continue;
            }

            throw new ParchmentRegistrationException(
                templateName,
                $"Block tags '{opening.Block!.Source}' and '{tag.Block!.Source}' are not in the same container. The paragraphs holding the tags of one block must be siblings: in the same table cell, the same text box, or at the same level of the body.",
                partUri,
                tag.Block.Source);
        }
    }

    static bool IsCloserOrAlternate(BlockTagKind kind, BlockTagKind closer) =>
        kind == closer ||
        (closer == BlockTagKind.EndIf && kind is BlockTagKind.ElsIf or BlockTagKind.Else);
}
