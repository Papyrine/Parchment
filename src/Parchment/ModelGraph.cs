/// <summary>
/// Traversal rules shared by the per-template maps, which each walk the model's type graph looking
/// for the members they dispatch.
/// </summary>
static class ModelGraph
{
    /// <summary>
    /// Whether a member's type is a composite worth walking into. Scalars, the framework types
    /// Parchment renders directly, and collections all end the walk.
    /// </summary>
    public static bool ShouldDescend(Type type)
    {
        if (type.IsPrimitive ||
            type.IsEnum)
        {
            return false;
        }

        if (type == typeof(string) ||
            type == typeof(decimal) ||
            type == typeof(DateTime) ||
            type == typeof(DateTimeOffset) ||
            type == typeof(Date) ||
            type == typeof(Time) ||
            type == typeof(TimeSpan) ||
            type == typeof(Guid) ||
            type == typeof(Uri))
        {
            return false;
        }

        return !typeof(IEnumerable).IsAssignableFrom(type);
    }
}
