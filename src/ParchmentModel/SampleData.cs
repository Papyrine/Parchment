namespace ParchmentSample;

public static class SampleData
{
    public static Invoice Invoice() =>
        new()
        {
            Number = "INV-2026-0042",
            IssueDate = new(2026, 4, 1),
            DueDate = new(2026, 4, 30),
            Currency = "USD",
            Notes = "Thanks for your business.",
            Tags = ["priority", "net-30", "repeat-customer"],
            Customer = new()
            {
                Name = "Globex Corporation",
                Email = "ap@globex.test",
                Address = "123 Commerce Way, Springfield, IL 62704",
                VatNumber = "US-123456789",
                IsPreferred = true
            },
            Lines =
            [
                new()
                {
                    Description = "Consulting services — API design",
                    Quantity = 12,
                    UnitPrice = 185m
                },
                new()
                {
                    Description = "Consulting services — implementation",
                    Quantity = 40,
                    UnitPrice = 165m
                },
                new()
                {
                    Description = "Project management",
                    Quantity = 8,
                    UnitPrice = 145m
                }
            ]
        };

    public static ReportContext Report() =>
        new()
        {
            Report = new()
            {
                Title = "Q2 Platform Health Review",
                Author = "Alex Chen",
                Date = new(2026, 7, 1),
                Summary = "Platform reliability and developer experience both trended upward this quarter, " +
                          "driven by the build pipeline rewrite and the new on-call rotation.",
                Findings =
                [
                    new() { Area = "Build", Status = "Improved", Owner = "DevEx" },
                    new() { Area = "Tests", Status = "Improved", Owner = "QA" },
                    new() { Area = "Deploys", Status = "Stable", Owner = "SRE" },
                    new() { Area = "On-call", Status = "Watch", Owner = "SRE" }
                ],
                Actions =
                [
                    new() { Title = "Roll out cached builds", Detail = "Extend the new pipeline to mobile repos." },
                    new() { Title = "Document runbooks", Detail = "Capture the top five paging scenarios." },
                    new() { Title = "Audit alert thresholds", Detail = "Reduce noise on the latency dashboard." }
                ],
                HasRisks = true
            }
        };
}
