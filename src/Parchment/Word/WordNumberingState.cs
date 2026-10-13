/// <summary>
/// Adapter for <see cref="NumberingDefinitionsPart"/> that reuses the existing definitions when present
/// and creates new ones on demand for bullet and ordered lists.
/// </summary>
class WordNumberingState(MainDocumentPart mainPart)
{
    int? bulletAbstractNumId;
    readonly Dictionary<NumberFormatValues, int> orderedAbstractNumIds = [];
    int nextAbstractNumId;
    int nextNumId;
    bool initialized;
    HtmlNumberingSession? htmlSession;

    internal HtmlNumberingSession GetHtmlSession() =>
        htmlSession ??= new();

    public int CreateBulletNumbering()
    {
        EnsureInitialized();
        var numbering = GetNumbering();
        var abstractId = bulletAbstractNumId ?? CreateBulletAbstract(numbering);
        bulletAbstractNumId = abstractId;
        return AppendInstance(numbering, abstractId, out _);
    }

    // The abstractNum is shared by every list of its format, and Word runs one count per level
    // through every instance of an abstractNum: a list carries on from the one before it unless its
    // own instance restarts the level. So every list states its start on its instance, a start of 1
    // included, on the level its items use. A nested list is an instance of its own, so that one
    // level is all an instance ever numbers.
    public int CreateOrderedNumbering(NumberFormatValues format, int ilvl = 0, int start = 1)
    {
        EnsureInitialized();
        var numbering = GetNumbering();
        if (!orderedAbstractNumIds.TryGetValue(format, out var abstractId))
        {
            abstractId = CreateOrderedAbstract(numbering, format);
            orderedAbstractNumIds[format] = abstractId;
        }

        var numId = AppendInstance(numbering, abstractId, out var instance);
        instance.Append(
            new LevelOverride
            {
                LevelIndex = ilvl,
                StartOverrideNumberingValue = new()
                {
                    Val = start
                }
            });

        return numId;
    }

    Numbering GetNumbering()
    {
        var part = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>(RelationshipIds.Named(mainPart, "rIdNumbering"));
        if (part.Numbering == null)
        {
            part.Numbering = new();
            part.Numbering.Save();
        }

        return part.Numbering;
    }

    void EnsureInitialized()
    {
        if (initialized)
        {
            return;
        }

        initialized = true;
        var part = mainPart.NumberingDefinitionsPart;
        if (part?.Numbering == null)
        {
            nextAbstractNumId = 1;
            nextNumId = 1;
            return;
        }

        foreach (var abstractNum in part.Numbering.Elements<AbstractNum>())
        {
            var idStr = abstractNum.AbstractNumberId?.Value;
            if (idStr >= nextAbstractNumId)
            {
                nextAbstractNumId = idStr.Value + 1;
            }
        }

        foreach (var numberingInstance in part.Numbering.Elements<NumberingInstance>())
        {
            var idStr = numberingInstance.NumberID?.Value;
            if (idStr >= nextNumId)
            {
                nextNumId = idStr.Value + 1;
            }
        }

        if (nextAbstractNumId == 0)
        {
            nextAbstractNumId = 1;
        }

        if (nextNumId == 0)
        {
            nextNumId = 1;
        }
    }

    // Word writes all nine levels for every abstractNum, and a paragraph whose ilvl has no matching
    // level renders unindented with no marker. The abstractNum is cached per format, so defining the
    // full set costs one definition rather than one per list.
    internal const int MaxIlvl = 8;

    // Where a level's text starts. A continuation paragraph of a list item carries no numPr, so it
    // gets none of the level's indentation and has to state this itself to line up under the item.
    internal static int TextIndent(int ilvl) =>
        480 * (ilvl + 1);

    // Word cycles these three glyphs across the nine levels.
    static readonly (string Glyph, string Font)[] bulletGlyphs =
    [
        ("\uF0B7", "Symbol"),
        ("o", "Courier New"),
        ("\uF0A7", "Wingdings")
    ];

    int CreateBulletAbstract(Numbering numbering)
    {
        var id = nextAbstractNumId++;
        var abstractNum = new AbstractNum { AbstractNumberId = id };
        for (var ilvl = 0; ilvl <= MaxIlvl; ilvl++)
        {
            var (glyph, font) = bulletGlyphs[ilvl % bulletGlyphs.Length];
            abstractNum.Append(BuildBulletLevel(ilvl, glyph, font));
        }

        InsertAbstract(numbering, abstractNum);
        return id;
    }

    int CreateOrderedAbstract(Numbering numbering, NumberFormatValues format)
    {
        var id = nextAbstractNumId++;
        var abstractNum = new AbstractNum { AbstractNumberId = id };
        for (var ilvl = 0; ilvl <= MaxIlvl; ilvl++)
        {
            abstractNum.Append(BuildOrderedLevel(ilvl, format));
        }

        InsertAbstract(numbering, abstractNum);
        return id;
    }

    // w:numbering is a sequence: every w:numPicBullet, then every w:abstractNum, then every w:num,
    // then w:numIdMacAtCleanup. Word refuses to open a document whose numbering part has an
    // abstractNum ahead of a picture bullet, so a new one follows the last of them.
    static void InsertAbstract(Numbering numbering, AbstractNum abstractNum)
    {
        var pictureBullet = numbering.Elements<NumberingPictureBullet>().LastOrDefault();
        if (pictureBullet == null)
        {
            numbering.InsertAt(abstractNum, 0);
            return;
        }

        numbering.InsertAfter(abstractNum, pictureBullet);
    }

    int AppendInstance(Numbering numbering, int abstractId, out NumberingInstance instance)
    {
        var numId = nextNumId++;
        instance = new()
        {
            NumberID = numId
        };
        instance.Append(
            new AbstractNumId
            {
                Val = abstractId
            });

        // The sequence closes with w:numIdMacAtCleanup, which Word for Mac writes after the
        // instances. A new instance goes in ahead of it, so it is not always the last child - hence
        // handing it back.
        if (numbering.LastChild is NumberingIdMacAtCleanup cleanup)
        {
            numbering.InsertBefore(instance, cleanup);
            return numId;
        }

        numbering.Append(instance);
        return numId;
    }

    static Level BuildBulletLevel(int ilvl, string glyph, string font) =>
        new()
        {
            LevelIndex = ilvl,
            NumberingFormat = new()
            {
                Val = NumberFormatValues.Bullet
            },
            LevelText = new()
            {
                Val = glyph
            },
            LevelJustification = new()
            {
                Val = LevelJustificationValues.Left
            },
            PreviousParagraphProperties = new(
                new Indentation
                {
                    Left = TextIndent(ilvl).ToString(),
                    Hanging = "240"
                }),
            NumberingSymbolRunProperties = new(
                new RunFonts
                {
                    Ascii = font,
                    HighAnsi = font,
                    Hint = FontTypeHintValues.Default
                })
        };

    static Level BuildOrderedLevel(int ilvl, NumberFormatValues format) =>
        new()
        {
            LevelIndex = ilvl,
            StartNumberingValue = new()
            {
                Val = 1
            },
            NumberingFormat = new()
            {
                Val = format
            },
            LevelText = new()
            {
                Val = $"%{ilvl + 1}."
            },
            LevelJustification = new()
            {
                Val = LevelJustificationValues.Left
            },
            PreviousParagraphProperties = new(
                new Indentation
                {
                    Left = TextIndent(ilvl).ToString(),
                    Hanging = "240"
                })
        };
}
