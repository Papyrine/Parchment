public class ShapeResolverTests
{
    static readonly ModelShape shape = BuildShape();

    [Test]
    public async Task Resolve_RootMember_ReturnsMemberType()
    {
        await Assert.That(ShapeResolver.TryResolve(shape, ["Customer"], emptyScope, out var result)).IsTrue();
        await Assert.That(result).IsEqualTo("global::Sample.Customer");
    }

    [Test]
    public async Task Resolve_NestedMember_WalksTypeChain()
    {
        await Assert.That(ShapeResolver.TryResolve(shape, ["Customer", "Name"], emptyScope, out var result)).IsTrue();
        await Assert.That(result).IsEqualTo("string");
    }

    [Test]
    public async Task Resolve_IsCaseInsensitive()
    {
        await Assert.That(ShapeResolver.TryResolve(shape, ["customer", "NAME"], emptyScope, out var result)).IsTrue();
        await Assert.That(result).IsEqualTo("string");
    }

    [Test]
    public async Task Resolve_UnknownMember_Fails() =>
        await Assert.That(ShapeResolver.TryResolve(shape, ["Customer", "DoesNotExist"], emptyScope, out _)).IsFalse();

    [Test]
    public async Task Resolve_UnknownRootMember_Fails() =>
        await Assert.That(ShapeResolver.TryResolve(shape, ["NotAField"], emptyScope, out _)).IsFalse();

    [Test]
    public async Task Resolve_EmptySegments_Fails() =>
        await Assert.That(ShapeResolver.TryResolve(shape, [], emptyScope, out _)).IsFalse();

    // "Customer.Name" resolves to string, but string isn't in the shape, so going further fails.
    [Test]
    public async Task Resolve_TraversingPrimitive_Fails() =>
        await Assert.That(ShapeResolver.TryResolve(shape, ["Customer", "Name", "Length"], emptyScope, out _)).IsFalse();

    [Test]
    public async Task Resolve_ScopedIdentifierShortCircuitsToBoundType()
    {
        // Loop variable `item` bound to Customer — "item.Name" should resolve via the binding,
        // not the root.
        var scope = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["item"] = "global::Sample.Customer"
        };

        await Assert.That(ShapeResolver.TryResolve(shape, ["item", "Name"], scope, out var result)).IsTrue();
        await Assert.That(result).IsEqualTo("string");
    }

    [Test]
    public async Task Resolve_ScopedIdentifierAlone_ReturnsBoundType()
    {
        var scope = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["item"] = "global::Sample.Customer"
        };

        await Assert.That(ShapeResolver.TryResolve(shape, ["item"], scope, out var result)).IsTrue();
        await Assert.That(result).IsEqualTo("global::Sample.Customer");
    }

    [Test]
    public async Task GetElementType_ReturnsConfiguredElement()
    {
        await Assert.That(ShapeResolver.TryGetElementType(shape, "global::Sample.Invoice", out var element)).IsTrue();
        await Assert.That(element).IsEqualTo("global::Sample.LineItem");
    }

    [Test]
    public async Task GetElementType_NonCollectionType_Fails() =>
        await Assert.That(ShapeResolver.TryGetElementType(shape, "global::Sample.Customer", out _)).IsFalse();

    [Test]
    public async Task GetElementType_UnknownType_Fails() =>
        await Assert.That(ShapeResolver.TryGetElementType(shape, "global::Sample.Unknown", out _)).IsFalse();

    // ReSharper disable once CollectionNeverUpdated.Local
    static Dictionary<string, string> emptyScope = new(StringComparer.OrdinalIgnoreCase);

    static ModelShape BuildShape()
    {
        // Sample.Invoice
        //   Customer : Sample.Customer
        //   Lines    : (collection of Sample.LineItem)
        // Sample.Customer
        //   Name     : string
        // Sample.LineItem
        //   Sku      : string
        var invoice = new TypeEntry(
            "global::Sample.Invoice",
            ElementTypeFullyQualifiedName: "global::Sample.LineItem",
            new(
            [
                new("Customer", "global::Sample.Customer"),
                new("Lines", "global::Sample.LineItemCollection")
            ]));

        var customer = new TypeEntry(
            "global::Sample.Customer",
            ElementTypeFullyQualifiedName: null,
            new([new("Name", "string")]));

        var lineItem = new TypeEntry(
            "global::Sample.LineItem",
            ElementTypeFullyQualifiedName: null,
            new([new("Sku", "string")]));

        return new(
            "global::Sample.Invoice",
            new([invoice, customer, lineItem]));
    }
}
