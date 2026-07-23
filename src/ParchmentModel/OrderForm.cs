namespace ParchmentSample;

#region GeneratorProtectionMode
[ParchmentModel("Templates/order-form.docx", Protection = ProtectionMode.None)]
public partial class OrderForm
{
    public required string Number { get; init; }

    [EditableField]
    public required string PurchaseOrder { get; set; }
}
#endregion
