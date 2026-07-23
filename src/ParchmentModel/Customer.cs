namespace ParchmentSample;

public class Customer
{
    public required string Name { get; init; }
    public required string Email { get; init; }
    public required string Address { get; init; }
    public string? VatNumber { get; init; }
    public bool IsPreferred { get; init; }
}
