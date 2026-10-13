/// <summary>
/// The attribute-map walk descends only into types that lead to a map member. These pin both
/// halves: a densely connected model with nothing to find is not walked path by path, and a member
/// behind a cycle, a struct or a collection element is still found.
/// </summary>
public class MapWalkPruningTests
{
    // Every type holds a member of every other type, so the simple paths through them number in
    // the factorials.
    static string Cluster(string prefix, int count)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < count; i++)
        {
            builder.Append($"public class {prefix}{i}\n{{\n    public string Text {{ get; set; }} = \"\";\n");
            for (var j = 0; j < count; j++)
            {
                if (j != i)
                {
                    builder.Append($"    public {prefix}{j} To{j} {{ get; set; }} = null!;\n");
                }
            }

            builder.Append("}\n");
        }

        return builder.ToString();
    }

    const string header =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute
        {
        }


        """;

    static string Generated(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

    static int Count(string text, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    // Walked path by path, twelve fully connected types take hours and terabytes of allocation, so
    // the timeout is what fails this.
    [Test]
    [Timeout(120_000)]
    public async Task DenseGraphWithoutMapMembers_IsNotWalked(Cancel cancel)
    {
        var source =
            header +
            """
            [ParchmentModel]
            public partial class Root
            {
                public string Name { get; set; } = "";
                public Node0 Start { get; set; } = null!;
            }

            """ +
            Cluster("Node", 12);

        var result = await Task.Run(() => GeneratorDriver.Run(source, "{{ Name }} {{ Start.To3.To7.Text }}"), cancel);

        await Assert.That(result.Diagnostics).IsEmpty();
        var generated = Generated(result);
        // Fluid accessors do not come from the walk: every type still gets its block.
        await Assert.That(generated).Contains("_Accessors_global__Sample_Node11");
        await Assert.That(generated).DoesNotContain("_Formats");
        await Assert.That(generated).DoesNotContain("_StringLists");
        await Assert.That(generated).DoesNotContain("_Editables");
    }

    // The marked members sit on Third, reached only by going round a cycle that leads back to
    // First, with a dense cluster holding nothing hanging off both the root and the cycle. A type
    // in a cycle must not be settled as leading nowhere before the rest of the cycle is known.
    [Test]
    [Timeout(120_000)]
    public async Task MapMembersBehindCycle_AreFoundOncePerPath(Cancel cancel)
    {
        var source =
            header +
            """
            [ParchmentModel]
            public partial class Root
            {
                public string Name { get; set; } = "";
                public Node0 Dense { get; set; } = null!;
                public First First { get; set; } = null!;
                public First Other { get; set; } = null!;

                [EditableField]
                public List<Item> Items { get; set; } = new();
            }

            public class First
            {
                public Third Back { get; set; } = null!;
                public Node0 Dense { get; set; } = null!;
                public Second Second { get; set; } = null!;
            }

            public class Second
            {
                public First Back { get; set; } = null!;
                public Third Third { get; set; } = null!;
            }

            public class Third
            {
                public First Back { get; set; } = null!;

                [Html]
                public string Body { get; set; } = "";

                public List<string> Tags { get; set; } = new();

                [EditableField]
                public string Note { get; set; } = "";
            }

            public class Item
            {
                [EditableField]
                public string Title { get; set; } = "";

                public Node0 Dense { get; set; } = null!;
                public Detail Detail { get; set; } = new();
            }

            public class Detail
            {
                [EditableField]
                public string Text { get; set; } = "";
            }

            """ +
            Cluster("Node", 12);

        var result = await Task.Run(() => GeneratorDriver.Run(source, "{{ Name }}"), cancel);

        await Assert.That(result.Diagnostics).IsEmpty();
        var generated = Generated(result);

        // First.Back reaches Third directly and First.Second.Third the long way; each from both
        // root members. A path may not revisit a type, so nothing longer exists.
        string[] prefixes = ["First.Back", "First.Second.Third", "Other.Back", "Other.Second.Third"];
        foreach (var prefix in prefixes)
        {
            await Assert.That(generated).Contains($"new(\"{prefix}.Body\", global::Parchment.Generated.FormatMapKind.Html");
            await Assert.That(generated).Contains($"\"{prefix}.Tags\"");
            await Assert.That(generated).Contains($"\"{prefix}.Note\"");
        }

        await Assert.That(Count(generated, "FormatMapKind.Html")).IsEqualTo(4);
        await Assert.That(Count(generated, ".Tags\"")).IsEqualTo(4);
        await Assert.That(Count(generated, ".Note\"")).IsEqualTo(4);

        // The element walk starts from the element type and is pruned the same way.
        await Assert.That(generated).Contains("\"Items\"");
        await Assert.That(generated).Contains("\"Title\"");
        await Assert.That(generated).Contains("\"Detail.Text\"");
    }

    // The struct is on the path, not the declaring type, and the hop below it carries no marker of
    // its own: the walk that reports PARCH032 has to get through it.
    [Test]
    [Timeout(120_000)]
    public async Task EditableBelowStruct_IsStillReported(Cancel cancel)
    {
        var source =
            header +
            """
            [ParchmentModel]
            public partial class Root
            {
                public string Name { get; set; } = "";
                public Node0 Dense { get; set; } = null!;
                public Wrap Wrap { get; set; }
            }

            public struct Wrap
            {
                public Middle Middle { get; set; }
            }

            public class Middle
            {
                public Leaf Leaf { get; set; } = null!;
            }

            public class Leaf
            {
                [EditableField]
                public string Deep { get; set; } = "";
            }

            """ +
            Cluster("Node", 12);

        var result = await Task.Run(() => GeneratorDriver.Run(source, "{{ Name }}"), cancel);

        await Assert.That(result.Diagnostics.Select(_ => _.Id)).Contains("PARCH032");
        await Assert.That(Generated(result)).DoesNotContain("\"Wrap.Middle.Leaf.Deep\"");
    }
}
