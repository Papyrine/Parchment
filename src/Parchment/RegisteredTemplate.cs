abstract class RegisteredTemplate(string name, Type modelType)
{
    public string Name { get; } = name;
    public Type ModelType { get; } = modelType;

    public abstract Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel);
}
