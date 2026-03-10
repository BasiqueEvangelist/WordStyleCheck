using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager
{
    private readonly List<ILint> _lints =
    [
        new PageSizeLint(),
        new PageMarginsLint(),
        new TocReferencesLint(),
        new HandmadeListLint(),
        new HandmadePageBreakLint(),
        new NeedlessParagraphLint(),
        new ForcePageBreakBeforeLint(x => x is {Class: ParagraphClass.Heading, HeadingData.Level: 1} or {Class: ParagraphClass.StructuralElementHeader}, "NeedsPageBreakBeforeHeader"),
        new ForceJustificationLint(x => x is {Class: ParagraphClass.StructuralElementHeader}, [JustificationValues.Center], "StructuralElementHeaderNotCentered"),
        new ForceJustificationLint(x => x is {CaptionData.Type: CaptionType.Table}, [JustificationValues.Left, JustificationValues.Both], "TableCaptionNotLeftAligned"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Appendix, OfNumbering: null}, 709, 0, "IncorrectBodyTextFirstLineIndent", "IncorrectBodyTextLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 0}, 709, 0, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 1}, -709, 1418, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 2}, -851, 1560, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphLineSpacingLint(
            // TODO: enforce this for numberings.
            // TODO: enforce this for table cell content, headers, captions.
            x => x is {Class: ParagraphClass.BodyText, OfNumbering: null, OfStructuralElement: not StructuralElement.Appendix, IsEmptyOrDrawing: false},
            360,
            "IncorrectTextLineSpacing"
        ),
        new ParagraphLineSpacingLint(
            // TODO: enforce this for numberings.
            // TODO: enforce this for table cell content, headers, captions.
            x => x is {Class: ParagraphClass.Caption},
            240,
            "IncorrectCaptionLineSpacing"
        ),
        new InterParagraphSpacingLint(
            [
                new(
                    x => x is {Class: ParagraphClass.Heading, OutlineLevel: 0},
                    0,
                    18 * 20
                ),
                new(
                    x => x is {Class: ParagraphClass.Heading, OutlineLevel: 1},
                    24 * 20,
                    12 * 20
                ),
                new(
                    x => x is {Class: ParagraphClass.Heading, OutlineLevel: 2},
                    12 * 20,
                    6 * 20
                ),
                new(
                    x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not (StructuralElement.Bibliography or StructuralElement.Appendix), OfNumbering: null} and not {OfStructuralElement: null, AssociatedHeading1: null},
                    6 * 20,
                    6 * 20,
                    contextualSpacing: true
                )
            ],
            "IncorrectInterParagraphSpacing"
        ),
        new CorrectStructuralElementHeaderLint(),
        new WrongCaptionPositionLint(CaptionType.Table, false, "IncorrectTableCaptionPosition"),
        new WrongCaptionPositionLint(CaptionType.Figure, true, "IncorrectFigureCaptionPosition"),
        new IncorrectCaptionTextLint(),
        new IncorrectCaptionedNumberingLint(CaptionType.Figure, "IncorrectFigureNumbering"),
        new IncorrectCaptionedNumberingLint(CaptionType.Table, "IncorrectTableNumbering"),
        new FigureTableNotReferencedLint(),
        new BibliographySourceNotReferencedLint(),
        new IncorrectHeadingTextLint(),
        // TODO: make this lint configurable.
        // new NotEnoughSourcesLint(40, "NotEnoughSources", "NoBibliography"),
        // new IncorrectOutlineLevelLint(x => x is { Class: ParagraphClass.BodyText }, _ => null, "BodyTextInToC"),
        // new IncorrectOutlineLevelLint(x => x is { HeadingData.Level: < 4 }, x => x.HeadingData!.Level - 1, "IncorrectHeaderOutlineLevel"),
        // new IncorrectOutlineLevelLint(x => x is { HeadingData.Level: 4 }, x => null, "SubPointsInToC"),
        // TODO: make text font lint configurable.
        // new TextFontLint(),
        new FontSizeLint(x => x is {Class: ParagraphClass.Heading or ParagraphClass.BodyText}, 24, "IncorrectFontSize"),
        new ForceBoldLint(true, x => x is { Class: ParagraphClass.Heading, OutlineLevel: null or < 2 } or {Class: ParagraphClass.StructuralElementHeader}, "HeadingNotBold"),
        new ForceBoldLint(false, x => x is { OutlineLevel: >= 2 }, "SubSubHeadingBold"),
        new ForceBoldLint(false, x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Bibliography}, "BodyTextBold"),
    ];

    public void Run(LintContext ctx)
    {
        foreach (var lint in _lints)
        {
            if (!lint.EmittedDiagnostics.Any(ctx.LintIdFilter.Invoke))
                continue;
            
            using (new LoudStopwatch(lint.GetType().Name))
            {
                try
                {
                    lint.Run(ctx);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Encountered exception while running {lint.GetType().Name}: {e}");
                }
            }
        }

        using (new LoudStopwatch("LintMerger.Run")) 
            LintMerger.Run(ctx.Messages);
    }

    public List<string> AllPossibleDiagnostics => _lints.SelectMany(x => x.EmittedDiagnostics).ToList();
}