/// <summary>
/// Warns when a render attribute sits on a static member. Dispatching one is a no-op: the
/// per-template maps behind <c>[ExcelsiorTable]</c>, <c>[Html]</c>, <c>[Markdown]</c> and
/// <c>[EditableField]</c> enumerate instance members only, so the value still binds and renders as
/// plain text while the attribute is dropped.
/// </summary>
/// <remarks>
/// The source generator reports the same thing as PARCH019 at compile time, but only for a model
/// carrying <c>[ParchmentModel]</c>. Hand registration never reaches the generator, so without this
/// the drop is entirely silent — the document renders, looks plausible, and is missing the table or
/// the html the attribute was there to produce.
/// <para>
/// Kept in step with <c>ParchmentTemplateGenerator.IgnoredStaticAttribute</c> and
/// <c>ShapeBuilder</c>: the same four attributes, and html/markdown detected through
/// <c>[StringSyntax]</c> as well as the user-defined attributes. Auto bullet-list dispatch on a
/// static <c>IEnumerable&lt;string&gt;</c> is deliberately excluded in both places, since it is
/// inferred from the member's type rather than written by hand.
/// </para>
/// </remarks>
static class StaticRenderAttributes
{
    public static void Warn(Type modelType, string templateName, ILogger logger)
    {
        var visited = new HashSet<Type>
        {
            modelType
        };
        WalkType(modelType, templateName, logger, visited);
    }

    static void WalkType(Type type, string templateName, ILogger logger, HashSet<Type> visited)
    {
        foreach (var member in EnumerateStatic(type))
        {
            var attribute = IgnoredAttribute(member);
            if (attribute == null)
            {
                continue;
            }

            logger.LogWarning(
                "Template {TemplateName}: [{Attribute}] on static member {Member} is ignored. The maps that dispatch it walk instance members only, so the value renders as plain text. Make the member an instance member to apply the attribute.",
                templateName,
                attribute,
                $"{type.Name}.{member.Name}");
        }

        // Descent mirrors the maps themselves, so a warning is only raised for a type they actually
        // walk into.
        foreach (var (_, memberType, _, _) in FormatMap.EnumerateMembers(type))
        {
            var underlying = Nullable.GetUnderlyingType(memberType) ?? memberType;
            if (!ModelGraph.ShouldDescend(underlying) ||
                !visited.Add(underlying))
            {
                continue;
            }

            WalkType(underlying, templateName, logger, visited);
            visited.Remove(underlying);
        }
    }

    static IEnumerable<MemberInfo> EnumerateStatic(Type type)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Static))
        {
            if (property.GetIndexParameters().Length > 0)
            {
                continue;
            }

            yield return property;
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Static))
        {
            yield return field;
        }
    }

    static string? IgnoredAttribute(MemberInfo member)
    {
        if (member.GetCustomAttribute<ExcelsiorTableAttribute>(true) != null)
        {
            return "ExcelsiorTable";
        }

        if (member.GetCustomAttribute<EditableFieldAttribute>(true) != null)
        {
            return "EditableField";
        }

        // [Html] and [Markdown] are user-defined, so they are matched by name exactly as
        // FormatMap.DetectFormat matches them.
        foreach (var attribute in member.GetCustomAttributes(true))
        {
            var name = attribute.GetType().Name;
            if (name == "HtmlAttribute")
            {
                return "Html";
            }

            if (name == "MarkdownAttribute")
            {
                return "Markdown";
            }
        }

        return member.GetCustomAttribute<StringSyntaxAttribute>(true)?.Syntax.ToLowerInvariant() switch
        {
            "html" => "StringSyntax(\"html\")",
            "markdown" => "StringSyntax(\"markdown\")",
            _ => null
        };
    }
}
