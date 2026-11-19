/// <summary>
/// The text a bound value is written as, in either flow.
/// </summary>
/// <remarks>
/// A number is the decimal as it is, in the invariant culture: <c>1234.50</c> keeps its scale and
/// <c>10.00</c> stays <c>10.00</c>. Fluid 3 formats a number the way Liquid prints one —
/// <c>1234.5</c>, and <c>10.0</c> — for its own output and for <c>ToStringValue</c> alike. Parchment
/// printed the scale before Fluid did that, and a currency amount losing its trailing zero is the
/// change a document would show. Anything else is Fluid's text for it.
/// </remarks>
static class FluidText
{
    public static string Of(FluidValue value, TemplateContext context)
    {
        if (value is NumberValue)
        {
            return value.ToNumberValue(context).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToStringValue(context);
    }
}
