/// <summary>
/// Makes the markdown flow print a number the way the docx flow does.
/// </summary>
/// <remarks>
/// The docx flow evaluates a token and takes <c>ToStringValue</c>, which prints the decimal as it
/// is: <c>1234.50</c> keeps its scale. The markdown flow renders the whole template, so the value is
/// written by Fluid's <c>NumberValue.WriteToAsync</c>, which reformats it — <c>1234.5</c>, and
/// <c>10.0</c> for <c>10.00</c>. Neither a value converter nor the encoder can reach that: the value
/// has to stay a number for comparisons and the math filters, and the encoder is handed text that
/// has already been reformatted. So each output expression is wrapped, and a number that reaches
/// the page is handed over as the text the docx flow would have written.
/// </remarks>
class MarkdownNumberText :
    AstRewriter
{
    static readonly MarkdownNumberText instance = new();

    public static IFluidTemplate Rewrite(IFluidTemplate template) =>
        instance.VisitTemplate(template);

    protected override Statement VisitOutputStatement(OutputStatement outputStatement) =>
        new OutputStatement(new NumberText(outputStatement.Expression));

    class NumberText(Expression inner) :
        Expression
    {
        public override ValueTask<FluidValue> EvaluateAsync(TemplateContext context)
        {
            var pending = inner.EvaluateAsync(context);
            if (pending.IsCompletedSuccessfully)
            {
                return new(Convert(pending.Result));
            }

            return Awaited(pending);
        }

        static async ValueTask<FluidValue> Awaited(ValueTask<FluidValue> pending) =>
            Convert(await pending);

        static FluidValue Convert(FluidValue value)
        {
            if (value is NumberValue)
            {
                return new Fluid.Values.StringValue(value.ToStringValue());
            }

            return value;
        }
    }
}
