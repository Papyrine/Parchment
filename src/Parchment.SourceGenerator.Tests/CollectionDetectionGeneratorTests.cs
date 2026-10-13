// Which member types the generator takes for a sequence, a dictionary or a dictionary entry. The
// types are recognised by name, arity and namespace, so the cases here are the ones a looser
// match would get wrong: a user type named after a collection, and a type that is one twice.
public class CollectionDetectionGeneratorTests
{
    const string model =
        """
        using System.Collections;
        using System.Collections.Generic;
        using Parchment;

        namespace Mine.Collections.Generic
        {
            public interface IEnumerable<T>
            {
                T Only { get; }
            }
        }

        namespace Sample
        {
            public class Line
            {
                public string Description { get; set; } = "";
            }

            public class Address
            {
                public string Street { get; set; } = "";
            }

            public class Twice : IEnumerable<Line>, IEnumerable<Address>
            {
                IEnumerator<Line> IEnumerable<Line>.GetEnumerator() => null!;
                IEnumerator<Address> IEnumerable<Address>.GetEnumerator() => null!;
                IEnumerator IEnumerable.GetEnumerator() => null!;
            }

            public class List<T>
            {
                public T Only { get; set; } = default!;
            }

            public class KeyValuePair<TKey, TValue>
            {
                public TValue Only { get; set; } = default!;
            }

            public interface IDictionary<TKey, TValue>
            {
                TValue Only { get; }
            }

            public class Holder
            {
                public interface IEnumerable<T>
                {
                    T Only { get; }
                }
            }

            public class Lookalike : Mine.Collections.Generic.IEnumerable<Line>, Holder.IEnumerable<Line>
            {
                public Line Only { get; set; } = new();
            }

            [ParchmentModel]
            public partial class Model
            {
                public string Name { get; set; } = "";
                public Line[] Array { get; set; } = [];
                public System.Collections.Generic.List<Line> Lines { get; set; } = [];
                public IEnumerable<Line> Sequence { get; set; } = [];
                public HashSet<Line> Set { get; set; } = [];
                public Twice Twice { get; set; } = new();
                public Dictionary<string, Line> Map { get; set; } = new();
                public System.Collections.Generic.IDictionary<string, Line> MapInterface { get; set; } = null!;
                public IReadOnlyDictionary<string, Line> ReadMap { get; set; } = null!;
                public IAsyncEnumerable<Line> Async { get; set; } = null!;
                public List<Line> NamedList { get; set; } = new();
                public KeyValuePair<string, Line> NamedPair { get; set; } = new();
                public IDictionary<string, Line> NamedMap { get; set; } = null!;
                public Lookalike Lookalike { get; set; } = new();
            }
        }
        """;

    [Test]
    [Arguments("Array")]
    [Arguments("Lines")]
    [Arguments("Sequence")]
    [Arguments("Set")]
    [Arguments("Twice")]
    public async Task SequenceYieldsItsElement(string source)
    {
        await Assert.That(Loop(source, "item.Description")).IsEmpty();
        await Assert.That(Loop(source, "item.Street")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    [Arguments("Map")]
    [Arguments("MapInterface")]
    [Arguments("ReadMap")]
    public async Task DictionaryYieldsItsEntries(string source)
    {
        await Assert.That(Loop(source, "item.Key")).IsEmpty();
        await Assert.That(Loop(source, "item.Value.Description")).IsEmpty();
        await Assert.That(Loop(source, "item.Value.Street")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Loop(source, "item.Only")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    [Arguments("Name")]
    [Arguments("Async")]
    [Arguments("NamedList")]
    [Arguments("NamedPair")]
    [Arguments("NamedMap")]
    [Arguments("Lookalike")]
    public async Task NotASequence(string source)
    {
        // The docx flow goes on to report the loop variable it could not type.
        await Assert.That(Ids(GeneratorDriver.Run(model, $"{{% for item in {source} %}}", "{{ item }}", "{% endfor %}"))).Contains("PARCH002");
        await Assert.That(Ids(GeneratorDriver.RunMarkdown(model, $"{{% for item in {source} %}}\n{{{{ item }}}}\n\n{{% endfor %}}"))).IsEquivalentTo(["PARCH002"]);
    }

    [Test]
    [Arguments("NamedList")]
    [Arguments("NamedPair")]
    [Arguments("NamedMap")]
    [Arguments("Lookalike")]
    public async Task NamesakeBindsItsOwnMembers(string member)
    {
        await Assert.That(Substitution($"{member}.Only.Description")).IsEmpty();
        await Assert.That(Substitution($"{member}.Key")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Substitution($"{member}.size")).IsEquivalentTo(["PARCH001"]);
    }

    static List<string> Loop(string source, string path)
    {
        var docx = Ids(GeneratorDriver.Run(model, $"{{% for item in {source} %}}", $"{{{{ {path} }}}}", "{% endfor %}"));
        var markdown = Ids(GeneratorDriver.RunMarkdown(model, $"{{% for item in {source} %}}\n{{{{ {path} }}}}\n\n{{% endfor %}}"));
        return Agreed(docx, markdown);
    }

    static List<string> Substitution(string path)
    {
        var docx = Ids(GeneratorDriver.Run(model, $"{{{{ {path} }}}}"));
        var markdown = Ids(GeneratorDriver.RunMarkdown(model, $"{{{{ {path} }}}}"));
        return Agreed(docx, markdown);
    }

    // Both flows have to give the one verdict; a flow that disagrees shows up as its ids appended.
    static List<string> Agreed(List<string> docx, List<string> markdown)
    {
        if (docx.SequenceEqual(markdown))
        {
            return docx;
        }

        return [..docx, "markdown:", ..markdown];
    }

    // PARCH012 is the missing removePersonalInformation warning every test docx trips.
    static List<string> Ids(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _ != "PARCH012")
            .ToList();
}
