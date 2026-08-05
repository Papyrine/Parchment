// ReSharper disable PartialTypeWithSinglePart
public partial class EnumSubstitutionTests
{
    [Test]
    public async Task PascalCaseEnum_IsHumanized()
    {
        // Inline `{{ Status }}` tokens used to render the raw CLR symbol (`FullTime`) because
        // Fluid's default type dispatch falls through to Enum.ToString(). The Excelsior
        // EnumRender ValueConverter routes enums through ValueRenderer / Humanize so the
        // displayed text matches what an [ExcelsiorTable] column emits for the same value.
        using var template = DocxTemplateBuilder.Build("Status: {{ Status }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EmploymentModel>(template);

        using var output = new MemoryStream();
        await store.Render(
            new EmploymentModel
            {
                Status = EmploymentStatus.FullTime
            },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task DisplayDescriptionAttribute_IsHonoured()
    {
        using var template = DocxTemplateBuilder.Build("Role: {{ Role }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<RoleModel>(template);

        using var output = new MemoryStream();
        await store.Render(
            new RoleModel
            {
                Role = Role.Architect
            },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task NullableEnum_WithValue_Humanized()
    {
        using var template = DocxTemplateBuilder.Build("Backup: {{ Backup }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<NullableEnumModel>(template);

        using var output = new MemoryStream();
        await store.Render(
            new NullableEnumModel
            {
                Backup = EmploymentStatus.PartTime
            },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task NullableEnum_Null_RendersEmpty()
    {
        using var template = DocxTemplateBuilder.Build("Backup: {{ Backup }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<NullableEnumModel>(template);

        using var output = new MemoryStream();
        await store.Render(
            new NullableEnumModel
            {
                Backup = null
            },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task TypedSetOverride_IsHonoured()
    {
        // EnumRender<TEnum>.Set is what the Excelsior source generator emits per enum.
        // Verifies the converter walks the full chain — not just the global override and
        // not just Humanize.
        EnumRender<OverrideEnum>.Set(static value => value switch
        {
            OverrideEnum.Alpha => "ALPHA-OVERRIDE",
            OverrideEnum.Bravo => "BRAVO-OVERRIDE",
            _ => value.ToString()
        });

        using var template = DocxTemplateBuilder.Build("Tier: {{ Tier }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<OverrideModel>(template);

        using var output = new MemoryStream();
        await store.Render(
            new OverrideModel
            {
                Tier = OverrideEnum.Bravo
            },
            output);
        await Verify(output, "docx");
    }

    [ParchmentBindable]
    public partial class EmploymentModel
    {
        public required EmploymentStatus Status { get; init; }
    }

    [ParchmentBindable]
    public partial class NullableEnumModel
    {
        public required EmploymentStatus? Backup { get; init; }
    }

    public enum EmploymentStatus
    {
        FullTime,
        PartTime,
        Contract
    }

    [ParchmentBindable]
    public partial class RoleModel
    {
        public required Role Role { get; init; }
    }

    public enum Role
    {
        [Display(Description = "Solutions Architect")]
        Architect,

        [Display(Name = "Senior Dev")]
        Developer
    }

    [ParchmentBindable]
    public partial class OverrideModel
    {
        public required OverrideEnum Tier { get; init; }
    }

    public enum OverrideEnum
    {
        Alpha,
        Bravo
    }
}
