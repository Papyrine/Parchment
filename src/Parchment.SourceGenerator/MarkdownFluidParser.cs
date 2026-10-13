using System.Text.Encodings.Web;
using System.Threading.Tasks;
using static Parlot.Fluent.Parsers;

/// <summary>
/// Fluid's parser with <c>{% liquid %}</c> replaced by a statement the validators can see.
/// </summary>
/// <remarks>
/// Fluid registers the tag and then does not run it. A single statement
/// (<c>{% liquid echo Name %}</c>) parses to no statement at all and takes the rest of the template
/// with it: <c>[{% liquid echo Name %}]</c> renders <c>[</c>. More than one statement, or a
/// <c>for</c> / <c>if</c> inside, is a parse error naming a line of the tag's body. Nothing is left
/// in the tree for a validator to find, so the tag is swapped for one that reads to its own
/// <c>%}</c> and leaves an <see cref="UnsupportedLiquidTagStatement"/> behind, which
/// <c>MarkdownValidator</c> rejects. Being a tag, it is not seen inside
/// <c>{% raw %}</c> or <c>{% comment %}</c>, where the text is not liquid.
/// A copy of the one in the runtime, which <c>MarkdownReferenceValidator</c> reads.
/// </remarks>
class MarkdownFluidParser :
    FluidParser
{
    public MarkdownFluidParser() =>
        RegisteredTags["liquid"] = AnyCharBefore(TagEnd, canBeEmpty: true)
            .SkipAnd(TagEnd)
            .Then<Statement>(static _ => new UnsupportedLiquidTagStatement());
}

class UnsupportedLiquidTagStatement :
    Statement
{
    public override ValueTask<Completion> WriteToAsync(TextWriter writer, TextEncoder encoder, TemplateContext context) =>
        throw new NotSupportedException("'{% liquid %}' is rejected when the template is registered and never renders.");
}
