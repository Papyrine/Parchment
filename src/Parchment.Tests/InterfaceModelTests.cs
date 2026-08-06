public class InterfaceModelTests
{
    public interface IModel
    {
        string Name { get; }
    }

    [Test]
    public async Task RegisterDocx_InterfaceModel_FailsWithGuidance()
    {
        using var template = DocxTemplateBuilder.Build("{{ Name }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<IModel>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("interface");
        await Assert.That(exception.Message).Contains("IModel");
    }

    [Test]
    public async Task RegisterMarkdown_InterfaceModel_FailsWithGuidance()
    {
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterMarkdownTemplate<IModel>("{{ Name }}"))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("interface");
        await Assert.That(exception.Message).Contains("IModel");
    }
}
