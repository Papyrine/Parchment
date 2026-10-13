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

    // nextNumId when the session was last handed out. It moves with every list numbered here.
    int htmlSessionNumId;

    // A session carries OpenXmlHtml's next id from one conversion to the next, and it reads the
    // part again only when the part no longer ends with the element the session left there. That
    // says where the last definition was written, not whether one was, so it is not left to the
    // session to notice an id taken here: after one, the next conversion is handed a session that
    // counts for itself. It keeps the bullet definition, which is all a session knows that the part
    // does not say. ConvertHtml is the same care taken in the other direction.
    internal HtmlNumberingSession GetHtmlSession()
    {
        if (htmlSession == null)
        {
            htmlSession = new();
        }
        else if (htmlSessionNumId != nextNumId)
        {
            htmlSession = new()
            {
                BulletAbstractNumId = htmlSession.BulletAbstractNumId
            };
        }

        htmlSessionNumId = nextNumId;
        return htmlSession;
    }

    /// <summary>
    /// Converts html, whose lists are numbered by OpenXmlHtml rather than by this class.
    /// </summary>
    /// <remarks>
    /// OpenXmlHtml finds a free id by reading the numbering part and writes its definitions straight
    /// into it. The counters here are read from the part once and counted on from there, so they
    /// know nothing of those: the next markdown list was handed an id an html list already held, and
    /// Word drew both lists from whichever definition it found first - an html ol with bullets, a
    /// markdown bullet numbered "3.". So every conversion that can hold a list comes through here,
    /// and one that wrote a list has the counters read again before the next definition is created.
    /// Whether it wrote one is told from the last w:num: a list needs an instance of its own, and an
    /// instance is written after the ones already there. That keeps html holding no list from
    /// costing a read of the part, and a render that converts no html still reads it once however
    /// many lists it holds.
    /// </remarks>
    public List<OpenXmlElement> ConvertHtml(string html, HtmlConvertSettings settings)
    {
        var last = LastInstance();
        // The converter adds the relationships of links and images to the main part itself.
        RelationshipIds.Forget(mainPart);
        var elements = WordHtmlConverter.ToElements(html, mainPart, settings);
        if (!ReferenceEquals(last, LastInstance()))
        {
            initialized = false;
        }

        return elements;
    }

    /// <summary>
    /// Converts html met on the markdown side: an html block, an image, a token value in a markdown
    /// template.
    /// </summary>
    /// <remarks>
    /// Without a session OpenXmlHtml reads every definition and instance in the part to find a free
    /// id, on every conversion, so a document of N html blocks read 1 + 2 + ... + N of them. The
    /// session carries the id across. It would also carry the bullet definition, which these
    /// conversions have never shared - each html block holding a ul writes one of its own - so that
    /// is withheld for the call and put back after it, leaving the part as it was written before
    /// the session was passed and the docx flow's own html tokens still sharing theirs.
    /// </remarks>
    public List<OpenXmlElement> ConvertHtml(string html, ImagePolicies imagePolicies, int headingOffset = 0)
    {
        var session = GetHtmlSession();
        var bullet = session.BulletAbstractNumId;
        session.BulletAbstractNumId = null;
        try
        {
            return ConvertHtml(html, imagePolicies.BuildSettings(headingOffset, session));
        }
        finally
        {
            session.BulletAbstractNumId = bullet;
        }
    }

    // Caller code has been handed the main part and can have written definitions anywhere in the
    // numbering part, which a session only notices at its end. The next conversion counts again.
    internal void ForgetHtmlIds()
    {
        if (htmlSession != null)
        {
            htmlSession = new()
            {
                BulletAbstractNumId = htmlSession.BulletAbstractNumId
            };
        }
    }

    // The element a new instance follows. All the schema allows after the instances is the
    // w:numIdMacAtCleanup Word for Mac writes. OpenXmlHtml appends past it, which makes its instance
    // the last child; one written in schema order would sit before it, so that is looked at too.
    // Stepping back walks the part, which only a part still ending that way pays for.
    OpenXmlElement? LastInstance()
    {
        var last = mainPart.NumberingDefinitionsPart?.Numbering?.LastChild;
        if (last is NumberingIdMacAtCleanup)
        {
            return last.PreviousSibling();
        }

        return last;
    }

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

        // One walk for both kinds of id. This runs again for a list that follows an html one - see
        // ConvertHtml - so a document alternating the two reads the part as often as it has lists.
        for (var child = part.Numbering.FirstChild; child != null; child = child.NextSibling())
        {
            switch (child)
            {
                case AbstractNum abstractNum when abstractNum.AbstractNumberId?.Value >= nextAbstractNumId:
                    nextAbstractNumId = abstractNum.AbstractNumberId.Value + 1;
                    break;
                case NumberingInstance numberingInstance when numberingInstance.NumberID?.Value >= nextNumId:
                    nextNumId = numberingInstance.NumberID.Value + 1;
                    break;
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
