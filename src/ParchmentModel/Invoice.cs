namespace ParchmentSample;

#region GeneratorDocxModel
[ParchmentModel("Templates/invoice.docx")]
public partial class Invoice
{
    public required string Number { get; init; }
    public required Date IssueDate { get; init; }
    public required Date DueDate { get; init; }
    public required Customer Customer { get; init; }
    public required IReadOnlyList<LineItem> Lines { get; init; }
    public required string Currency { get; init; }
    public string? Notes { get; init; }
    public IReadOnlyList<string> Tags { get; init; } = [];

    public decimal Subtotal =>
        Lines.Sum(x => x.LineTotal);

    public decimal Tax =>
        Math.Round(Subtotal * 0.1m, 2);

    public decimal Total =>
        Subtotal + Tax;
}
#endregion
