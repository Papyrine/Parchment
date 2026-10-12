// The runtime half of StructuralLockstepCases: every template in the list is refused at
// registration, with the message the list records. StructuralLockstepGeneratorTests holds the
// source generator to the same list, so a structural mistake is a build error and not a first
// render that throws.
public partial class StructuralLockstepTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required bool Flag { get; init; }
        public required string Name { get; init; }
        public required List<Line> Lines { get; init; }
    }

    public class Line
    {
        public required string Description { get; init; }
    }

    [Test]
    [MethodDataSource(typeof(StructuralLockstepCases), nameof(StructuralLockstepCases.Names))]
    public async Task RegistrationRejects(string name)
    {
        var structuralCase = StructuralLockstepCases.Get(name);
        using var template = DocxTemplateBuilder.Build(string.Join("\n\n", structuralCase.Paragraphs));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains(structuralCase.Runtime);
    }
}
