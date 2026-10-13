/// <summary>
/// Per-render allocator for the ids shared by editable-field elements: <c>w:sdt</c> ids and the
/// <c>w:permStart</c>/<c>w:permEnd</c> pair ids. Sequential in processing order, which keeps
/// renders deterministic (same template + same model → same ids) without any randomness.
/// </summary>
/// <remarks>
/// Both kinds of id have to be unique across the whole document, headers, footers and notes
/// included, and a template can carry content controls and editable ranges of its own. So the
/// counter is seeded at registration with the highest id of either kind in any part
/// (<see cref="MaxId"/>) and a render numbers from above it. Only positive ids are handed out.
/// </remarks>
sealed class EditableState(int max)
{
    int last = max;

    public int NextId() => ++last;

    /// <summary>
    /// The highest <c>w:sdt</c> or permission-range id already present in <paramref name="root"/>,
    /// or 0 when none is above that. Word writes ids anywhere in the signed 32-bit range; a
    /// negative one is never in the way of a counter that starts at 1.
    /// </summary>
    public static int MaxId(OpenXmlCompositeElement root)
    {
        var max = 0;
        foreach (var element in root.Descendants())
        {
            var id = element switch
            {
                SdtId sdt => sdt.Val?.InnerText,
                PermStart start => start.Id?.InnerText,
                PermEnd end => end.Id?.InnerText,
                _ => null
            };

            if (int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) &&
                value > max)
            {
                max = value;
            }
        }

        return max;
    }
}
