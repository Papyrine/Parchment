namespace ParchmentSample;

public class LineItem
{
    public required string Description { get; init; }
    public required int Quantity { get; init; }
    public required decimal UnitPrice { get; init; }

    public decimal LineTotal =>
        Quantity * UnitPrice;
}
