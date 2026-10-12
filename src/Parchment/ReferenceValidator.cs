[RequiresUnreferencedCode(Unsupported.Trimming)]
class ReferenceValidator(Type modelType, string templateName, string partUri)
{
    // Scope binding for a loop variable whose source is not statically known: a range, or a path
    // rooted at another such variable. The variable exists, but members reached through it are not
    // checked. Mirrors the `untyped` set in MarkdownReferenceValidator.
    static readonly Type untyped = typeof(void);

    public void ValidateTree(IReadOnlyList<RangeNode> nodes)
    {
        var scope = new Dictionary<string, Type>(StringComparer.Ordinal);
        WalkNodes(nodes, scope);
    }

    void WalkNodes(IReadOnlyList<RangeNode> nodes, Dictionary<string, Type> scope)
    {
        foreach (var node in nodes)
        {
            WalkNode(node, scope);
        }
    }

    void WalkNode(RangeNode node, Dictionary<string, Type> scope)
    {
        switch (node)
        {
            case SubstitutionNode substitutionNode:
                foreach (var token in substitutionNode.Tokens)
                {
                    ValidateReferences(token.References, scope, token.Source);
                }

                break;
            case LoopNode loop:
                var sourceRefs = IdentifierVisitor.Collect(loop.LoopSource);
                Type? elementType = null;
                if (IsUntypedSource(loop.LoopSource, sourceRefs, scope))
                {
                    // Nothing to resolve an element type from, which is an unknown, not a mistake.
                    // Members used as range bounds are still references like any other.
                    ValidateReferences(sourceRefs, scope, "for");
                    elementType = untyped;
                }
                else if (sourceRefs.Count > 0 &&
                         TryResolvePathType(sourceRefs[0], scope, out var iterableType))
                {
                    ModelValidator.TryResolveElementType(iterableType, out elementType);
                }

                if (elementType == null)
                {
                    var firstRef = sourceRefs.Count > 0 ? sourceRefs[0].ToString() : null;
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"Loop source '{firstRef}' does not resolve to an enumerable type.",
                        partUri,
                        loop.LoopVariable);
                }

                // The modifiers are evaluated before the first iteration, so they validate against
                // the enclosing scope: the loop variable is not bound yet.
                ValidateLoopModifier(loop.Limit, scope);
                ValidateLoopModifier(loop.Offset, scope);

                var loopScope = new Dictionary<string, Type>(scope, StringComparer.Ordinal)
                {
                    [loop.LoopVariable] = elementType
                };
                WalkNodes(loop.Body, loopScope);
                break;
            case IfNode ifNode:
                foreach (var branch in ifNode.Branches)
                {
                    ValidateReferences(IdentifierVisitor.Collect(branch.Condition), scope, "if");

                    WalkNodes(branch.Body, scope);
                }

                WalkNodes(ifNode.ElseBody, scope);
                break;
            case StaticNode:
                break;
        }
    }

    void ValidateReferences(IReadOnlyList<IdentifierPath> references, Dictionary<string, Type> scope, string tokenSource)
    {
        foreach (var reference in references)
        {
            if (IsUntyped(reference, scope))
            {
                continue;
            }

            ModelValidator.Validate(modelType, reference, scope, templateName, partUri, tokenSource);
        }
    }

    // Fluid parses a loop source as either a member path or a range. A range may name members in
    // its bounds, so the absence of references cannot be what identifies one.
    static bool IsUntypedSource(Expression source, IReadOnlyList<IdentifierPath> sourceRefs, Dictionary<string, Type> scope)
    {
        if (source is not MemberExpression)
        {
            return true;
        }

        return sourceRefs.Count > 0 &&
               IsUntyped(sourceRefs[0], scope);
    }

    static bool IsUntyped(IdentifierPath path, Dictionary<string, Type> scope) =>
        scope.TryGetValue(path.Root, out var bound) &&
        bound == untyped;

    void ValidateLoopModifier(Expression? modifier, Dictionary<string, Type> scope)
    {
        if (modifier == null)
        {
            return;
        }

        ValidateReferences(IdentifierVisitor.Collect(modifier), scope, "for");
    }

    bool TryResolvePathType(
        IdentifierPath path,
        Dictionary<string, Type> scope,
        [NotNullWhen(true)] out Type? pathType)
    {
        pathType = null;
        if (!scope.TryGetValue(path.Root, out var current) &&
            !ModelValidator.TryResolveMember(modelType, path.Root, out current))
        {
            return false;
        }

        for (var i = 1; i < path.Segments.Count; i++)
        {
            if (!ModelValidator.TryResolveSegment(current, path.Segments[i], out var next))
            {
                return false;
            }

            current = next;
        }

        pathType = current;
        return true;
    }
}
