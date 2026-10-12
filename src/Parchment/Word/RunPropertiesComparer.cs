/// <summary>
/// Decides whether two <c>w:rPr</c> would serialise to the same xml, without serialising either.
/// </summary>
/// <remarks>
/// <see cref="OpenXmlMarkdownRenderer.FlushParagraph"/> asks this of every adjacent pair of runs, and
/// <see cref="OpenXmlElement.OuterXml"/> builds a writer and a string each time it is read. The
/// answer has to be the one comparing <c>OuterXml</c> gives, since a pair that compares differently
/// is a pair of runs merged or left apart differently: a missing rPr is not an empty one, attributes
/// are compared in the order they are written, and anything that is not a typed element is still
/// compared by its xml.
/// </remarks>
static class RunPropertiesComparer
{
    public static bool Same(RunProperties? left, RunProperties? right)
    {
        if (left == null)
        {
            return right == null;
        }

        return right != null &&
               SameElement(left, right);
    }

    static bool SameElement(OpenXmlElement left, OpenXmlElement right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        // A typed element has one name and one prefix per type, so the type stands in for both.
        if (left.GetType() != right.GetType())
        {
            return false;
        }

        // An unknown element carries its own prefix and may hold text beside its children; a misc
        // node is a comment or a processing instruction. Neither is written by a renderer here.
        if (left is OpenXmlUnknownElement or OpenXmlMiscNode)
        {
            return left.OuterXml == right.OuterXml;
        }

        if (!SameAttributes(left, right) ||
            !SameMarkupCompatibility(left.MCAttributes, right.MCAttributes) ||
            !SameNamespaceDeclarations(left, right))
        {
            return false;
        }

        if (left is OpenXmlLeafTextElement leftText)
        {
            return leftText.Text == ((OpenXmlLeafTextElement) right).Text;
        }

        var leftChild = left.FirstChild;
        var rightChild = right.FirstChild;
        while (leftChild != null &&
               rightChild != null)
        {
            if (!SameElement(leftChild, rightChild))
            {
                return false;
            }

            leftChild = leftChild.NextSibling();
            rightChild = rightChild.NextSibling();
        }

        return leftChild == null &&
               rightChild == null;
    }

    // In order, because that is how they are written: schema order for the typed ones, which one
    // type always shares, then the extended ones as they were added.
    static bool SameAttributes(OpenXmlElement left, OpenXmlElement right)
    {
        if (!left.HasAttributes)
        {
            return !right.HasAttributes;
        }

        if (!right.HasAttributes)
        {
            return false;
        }

        var leftAttributes = left.GetAttributes();
        var rightAttributes = right.GetAttributes();
        if (leftAttributes.Count != rightAttributes.Count)
        {
            return false;
        }

        for (var i = 0; i < leftAttributes.Count; i++)
        {
            if (leftAttributes[i] != rightAttributes[i])
            {
                return false;
            }
        }

        return true;
    }

    static bool SameMarkupCompatibility(MarkupCompatibilityAttributes? left, MarkupCompatibilityAttributes? right)
    {
        if (left == null &&
            right == null)
        {
            return true;
        }

        return left?.Ignorable?.Value == right?.Ignorable?.Value &&
               left?.ProcessContent?.Value == right?.ProcessContent?.Value &&
               left?.PreserveElements?.Value == right?.PreserveElements?.Value &&
               left?.PreserveAttributes?.Value == right?.PreserveAttributes?.Value &&
               left?.MustUnderstand?.Value == right?.MustUnderstand?.Value;
    }

    static bool SameNamespaceDeclarations(OpenXmlElement left, OpenXmlElement right)
    {
        using var leftDeclarations = left.NamespaceDeclarations.GetEnumerator();
        using var rightDeclarations = right.NamespaceDeclarations.GetEnumerator();
        while (leftDeclarations.MoveNext())
        {
            if (!rightDeclarations.MoveNext() ||
                leftDeclarations.Current.Key != rightDeclarations.Current.Key ||
                leftDeclarations.Current.Value != rightDeclarations.Current.Value)
            {
                return false;
            }
        }

        return !rightDeclarations.MoveNext();
    }
}
