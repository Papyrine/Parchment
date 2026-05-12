/// <summary>
/// ImmutableArray wrapper with structural value equality, safe to flow through the
/// incremental pipeline so downstream stages can be cached when contents are unchanged.
/// </summary>
readonly struct EquatableArray<T>(ImmutableArray<T> array) :
    IEquatable<EquatableArray<T>>,
    IEnumerable<T>
// ReSharper disable once RedundantNotNullConstraint
    where T : notnull, IEquatable<T>
{
    public static readonly EquatableArray<T> Empty = new(ImmutableArray<T>.Empty);

    public int Count => array.IsDefault ? 0 : array.Length;

    public T this[int index] => array[index];

    public ImmutableArray<T> AsImmutableArray() =>
        array.IsDefault ? ImmutableArray<T>.Empty : array;

    public bool Equals(EquatableArray<T> other)
    {
        var a = AsImmutableArray();
        var b = other.AsImmutableArray();
        if (a.Length != b.Length)
        {
            return false;
        }

        for (var i = 0; i < a.Length; i++)
        {
            if (!a[i].Equals(b[i]))
            {
                return false;
            }
        }

        return true;
    }

    public override bool Equals(object? obj) =>
        obj is EquatableArray<T> other &&
        Equals(other);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (var item in AsImmutableArray())
        {
            hash.Add(item);
        }

        return hash.ToHashCode();
    }

    public IEnumerator<T> GetEnumerator() =>
        ((IEnumerable<T>) AsImmutableArray()).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    public static bool operator ==(EquatableArray<T> left, EquatableArray<T> right) =>
        left.Equals(right);

    public static bool operator !=(EquatableArray<T> left, EquatableArray<T> right) =>
        !left.Equals(right);
}
