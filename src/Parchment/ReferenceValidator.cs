class ReferenceValidator(Type modelType, string templateName, string partUri)
{
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
                    foreach (var reference in token.References)
                    {
                        ModelValidator.Validate(modelType, reference, scope, templateName, partUri, token.Source);
                    }
                }

                break;
            case LoopNode loop:
                var sourceRefs = IdentifierVisitor.Collect(loop.LoopSource);
                Type? elementType = null;
                if (sourceRefs.Count > 0)
                {
                    var iterableType = ResolvePathType(sourceRefs[0], scope);
                    if (iterableType != null)
                    {
                        ModelValidator.TryResolveElementType(iterableType, out elementType);
                    }
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

                var loopScope = new Dictionary<string, Type>(scope, StringComparer.Ordinal)
                {
                    [loop.LoopVariable] = elementType
                };
                WalkNodes(loop.Body, loopScope);
                break;
            case IfNode ifNode:
                foreach (var branch in ifNode.Branches)
                {
                    var branchRefs = IdentifierVisitor.Collect(branch.Condition);
                    foreach (var reference in branchRefs)
                    {
                        ModelValidator.Validate(modelType, reference, scope, templateName, partUri, "if");
                    }

                    WalkNodes(branch.Body, scope);
                }

                WalkNodes(ifNode.ElseBody, scope);
                break;
            case StaticNode:
                break;
        }
    }

    Type? ResolvePathType(IdentifierPath path, Dictionary<string, Type> scope)
    {
        if (!scope.TryGetValue(path.Root, out var current) &&
            !ModelValidator.TryResolveMember(modelType, path.Root, out current))
        {
            return null;
        }

        for (var i = 1; i < path.Segments.Count; i++)
        {
            if (!ModelValidator.TryResolveMember(current, path.Segments[i], out var next))
            {
                return null;
            }

            current = next;
        }

        return current;
    }
}
