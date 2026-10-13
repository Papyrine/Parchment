/// <summary>
/// The models of one compilation share the entry of every type more than one of them reaches
/// (<c>ShapeBuilder.sharedEntries</c>). What a model is given must be what it would have built
/// alone, so each test generates the models together and one at a time and compares the files.
/// </summary>
public class SharedShapeGeneratorTests
{
    const string prelude =
        """
        using System;
        using System.Collections.Generic;
        using Parchment;

        namespace Excelsior
        {
            public class WordTableBuilder<T>;
        }

        """;

    static Dictionary<string, string> Generate(string source, params string[] models)
    {
        foreach (var model in models)
        {
            source = source.Replace($"/*{model}*/", "[ParchmentBindable]");
        }

        var setup = GeneratorDriver.CreateDriverWithFiles(prelude + source);
        return Sources(setup.Driver.RunGenerators(setup.Compilation));
    }

    static Dictionary<string, string> Sources(Microsoft.CodeAnalysis.GeneratorDriver driver) =>
        driver.GetRunResult()
            .Results.Single()
            .GeneratedSources
            .ToDictionary(_ => _.HintName, _ => _.SourceText.ToString());

    // Both declaration orders, since the model declared first is the one whose walk fills the
    // shared entries the other then reads.
    static async Task<(string first, string second)> AssertSameAsAlone(string firstModel, string secondModel, string firstSource, string secondSource, string sharedSource)
    {
        var firstAlone = Generate(firstSource + sharedSource, firstModel);
        var secondAlone = Generate(secondSource + sharedSource, secondModel);
        foreach (var source in new[] {firstSource + secondSource + sharedSource, secondSource + firstSource + sharedSource})
        {
            var together = Generate(source, firstModel, secondModel);
            await Assert.That(together.Count).IsEqualTo(firstAlone.Count + secondAlone.Count);
            foreach (var pair in firstAlone.Concat(secondAlone))
            {
                await Assert.That(together[pair.Key]).IsEqualTo(pair.Value);
            }
        }

        return (Join(firstAlone), Join(secondAlone));
    }

    // A model's files as one text: the forwarders to a Configure method are a file of their own.
    static string Join(Dictionary<string, string> sources) =>
        string.Concat(sources.OrderBy(_ => _.Key, StringComparer.Ordinal).Select(_ => _.Value));

    // The order of a shape's types is the order the model's own walk meets them in, and the
    // pragma carries only the ids that model's graph raises.
    [Test]
    public async Task EachModelKeepsItsOwnWalkOrderAndSuppressions()
    {
        var (invoice, receipt) = await AssertSameAsAlone(
            "Invoice",
            "Receipt",
            """
            namespace Sample
            {
                /*Invoice*/
                public partial class Invoice
                {
                    public Customer Customer { get; set; } = new();
                    public List<Line> Lines { get; set; } = [];
                    public Legacy Legacy { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                /*Receipt*/
                public partial class Receipt
                {
                    public Dictionary<string, Line> ByCode { get; set; } = [];
                    public Customer Customer { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                public class Customer
                {
                    public string Name { get; set; } = "";
                    public Address Address { get; set; } = new();
                }

                public class Address
                {
                    public string City { get; set; } = "";
                    public Customer? Occupant { get; set; }
                }

                public class Line
                {
                    public string Description { get; set; } = "";
                    public Customer? Buyer { get; set; }
                }

                public class Legacy
                {
                    [Obsolete("old")]
                    public string Code { get; set; } = "";
                }
            }
            """);

        await Assert.That(invoice).Contains("CS0618");
        await Assert.That(receipt).DoesNotContain("CS0618");
    }

    // A private member takes its name from the base types only for code that can see it: the
    // model nested in Derived, not the one beside it. That entry is each model's own.
    [Test]
    public async Task PrivateMemberHidesOnlyForTheModelThatSeesIt()
    {
        var (outer, inner) = await AssertSameAsAlone(
            "Outer",
            "Inner",
            """
            namespace Sample
            {
                /*Outer*/
                public partial class Outer
                {
                    public Derived Derived { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                public partial class Derived
                {
                    /*Inner*/
                    public partial class Inner
                    {
                        public Derived Derived { get; set; } = new();
                    }
                }
            }

            """,
            """
            namespace Sample
            {
                public class Base
                {
                    public string Hidden { get; set; } = "";
                }

                public partial class Derived : Base
                {
                    public string Shown { get; set; } = "";
                #pragma warning disable CS0169
                    new int Hidden;
                #pragma warning restore CS0169
                }
            }
            """);

        await Assert.That(outer).Contains("\"Hidden\"");
        await Assert.That(inner).DoesNotContain("\"Hidden\"");
        await Assert.That(inner).Contains("\"Shown\"");
    }

    // The two members' types print the same, and the walk tells types apart by how they print,
    // but the nullability of Box<T>.Value is read off the symbol.
    [Test]
    public async Task NullabilityOfATypeArgumentIsNotShared()
    {
        var (plain, annotated) = await AssertSameAsAlone(
            "Plain",
            "Annotated",
            """
            namespace Sample
            {
                /*Plain*/
                public partial class Plain
                {
                    public Box<string> Box { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                /*Annotated*/
                public partial class Annotated
                {
                    public Box<string?> Box { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                public class Box<T>
                {
                    [EditableField]
                    public T Value { get; set; } = default!;
                }
            }
            """);

        await Assert.That(annotated.Replace("Annotated", "Plain")).IsNotEqualTo(plain);
    }

    // A Configure method out of the model's reach is called through a forwarder named after the
    // model, so the member naming it is resolved per model.
    [Test]
    public async Task ConfigureForwarderIsNamedAfterItsOwnModel()
    {
        var (first, second) = await AssertSameAsAlone(
            "First",
            "Second",
            """
            namespace Sample
            {
                /*First*/
                public partial class First
                {
                    public Report Report { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                /*Second*/
                public partial class Second
                {
                    public Report Report { get; set; } = new();
                }
            }

            """,
            """
            namespace Sample
            {
                public class Line
                {
                    public string Description { get; set; } = "";
                }

                public partial class Report
                {
                    [ExcelsiorTable(Configure = nameof(Setup))]
                    public List<Line> Lines { get; set; } = [];

                    static void Setup(Excelsior.WordTableBuilder<Line> builder)
                    {
                    }
                }
            }
            """);

        await Assert.That(first).Contains("ParchmentConfigure_Sample_First_Setup");
        await Assert.That(first).DoesNotContain("Sample_Second");
        await Assert.That(second).Contains("ParchmentConfigure_Sample_Second_Setup");
        await Assert.That(second).DoesNotContain("Sample_First");
    }

    static string BuildGraph(int types, int models)
    {
        var builder = new StringBuilder("namespace Sample\n{\n");
        for (var index = 0; index < types; index++)
        {
            builder.Append($"public class T{index}\n{{\npublic string Name {{ get; set; }} = \"\";\npublic int? Count {{ get; set; }}\n");
            if (index + 1 < types)
            {
                builder.Append($"public T{index + 1} Next {{ get; set; }} = null!;\n");
            }

            if (index * 2 + 2 < types)
            {
                builder.Append($"public List<T{index * 2 + 2}> Items {{ get; set; }} = [];\n");
            }

            builder.Append("}\n");
        }

        for (var index = 0; index < models; index++)
        {
            builder.Append($"[ParchmentBindable]\npublic partial class M{index}\n{{\npublic T{index % 5} Root {{ get; set; }} = null!;\n}}\n");
        }

        builder.Append("}\n");
        return builder.ToString();
    }

    // The shared entries are filled by whichever run of the generator gets to a type first, and a
    // host may run the generator over one compilation from several threads.
    [Test]
    public async Task ConcurrentRunsOverOneCompilationAgree()
    {
        var source = prelude + BuildGraph(40, 12);
        var expected = Sources(Run(GeneratorDriver.CreateDriverWithFiles(source)));
        await Assert.That(expected.Count).IsEqualTo(12);

        for (var round = 0; round < 5; round++)
        {
            // A compilation of its own each round, so every round starts with nothing shared.
            var setup = GeneratorDriver.CreateDriverWithFiles(source);
            var results = new Dictionary<string, string>[16];
            using var barrier = new Barrier(results.Length);
            var threads = new Thread[results.Length];
            for (var index = 0; index < threads.Length; index++)
            {
                var slot = index;
                threads[index] = new(() =>
                {
                    barrier.SignalAndWait();
                    results[slot] = Sources(Run(setup));
                });
                threads[index].Start();
            }

            foreach (var thread in threads)
            {
                thread.Join();
            }

            foreach (var result in results)
            {
                await Assert.That(result.Count).IsEqualTo(expected.Count);
                foreach (var pair in expected)
                {
                    await Assert.That(result[pair.Key]).IsEqualTo(pair.Value);
                }
            }
        }
    }

    static Microsoft.CodeAnalysis.GeneratorDriver Run(GeneratorDriver.DriverSetup setup) =>
        setup.Driver.RunGenerators(setup.Compilation);
}
