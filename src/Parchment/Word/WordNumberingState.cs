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
        return AppendInstance(numbering, abstractId);
    }

    public int CreateOrderedNumbering(NumberFormatValues format)
    {
        EnsureInitialized();
        var numbering = GetNumbering();
        if (!orderedAbstractNumIds.TryGetValue(format, out var abstractId))
        {
            abstractId = CreateOrderedAbstract(numbering, format);
            orderedAbstractNumIds[format] = abstractId;
        }

        return AppendInstance(numbering, abstractId);
    }

    Numbering GetNumbering()
    {
        var part = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
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

        numbering.InsertAt(abstractNum, 0);
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

        numbering.InsertAt(abstractNum, 0);
        return id;
    }

    int AppendInstance(Numbering numbering, int abstractId)
    {
        var numId = nextNumId++;
        var instance = new NumberingInstance
        {
            NumberID = numId
        };
        instance.Append(
            new AbstractNumId
            {
                Val = abstractId
            });
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
                    Left = (480 * (ilvl + 1)).ToString(),
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
                    Left = (480 * (ilvl + 1)).ToString(),
                    Hanging = "240"
                })
        };
}
