namespace ParchmentSample;

#region ReportModel
#region GeneratorMarkdownModel
[ParchmentModel("Templates/report.md")]
public partial class ReportContext
{
    public required Report Report;
}
#endregion

public class Report
{
    public required string Title;
    public required string Author;
    public required Date Date;
    public required string Summary;
    public required IReadOnlyList<Finding> Findings;
    public required IReadOnlyList<ActionItem> Actions;
    public required bool HasRisks;
}

public class Finding
{
    public required string Area;
    public required string Status;
    public required string Owner;
}

public class ActionItem
{
    public required string Title;
    public required string Detail;
}
#endregion
