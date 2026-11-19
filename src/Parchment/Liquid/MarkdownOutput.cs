/// <summary>
/// The output the markdown flow renders liquid into: everything written lands in
/// <see cref="Builder"/> at once.
/// </summary>
/// <remarks>
/// Fluid's own outputs buffer and reach their destination only when flushed. Two things read this
/// text while the render is still writing it: <see cref="MarkdownEncoder"/>, which escapes a value
/// by where it lands, and <see cref="MarkdownBlankOutputs"/>, which records where each blank value
/// went. Both need every character written so far to be in the builder.
/// </remarks>
sealed class MarkdownOutput :
    IFluidOutput
{
    char[] scratch = [];

    public StringBuilder Builder { get; } = new();

    public void Write(string value) =>
        Builder.Append(value);

    public void Write(char[] buffer, int index, int count) =>
        Builder.Append(buffer, index, count);

    public ValueTask FlushAsync() =>
        default;

    // Fluid encodes into the span it is handed and then advances by what it wrote. The span is
    // scratch space; Advance is what appends it.
    public Span<char> GetSpan(int sizeHint = 0) =>
        Scratch(sizeHint);

    public Memory<char> GetMemory(int sizeHint = 0) =>
        Scratch(sizeHint);

    public void Advance(int count) =>
        Builder.Append(scratch, 0, count);

    char[] Scratch(int sizeHint)
    {
        var size = Math.Max(sizeHint, 256);
        if (scratch.Length < size)
        {
            scratch = new char[Math.Max(size, scratch.Length * 2)];
        }

        return scratch;
    }
}
