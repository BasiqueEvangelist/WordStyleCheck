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
        new HandmadeListLint(),
        new HandmadePageBreakLint(),
        new NeedlessParagraphLint(),
        new ForceJustificationLint(x => x is {Class: ParagraphClass.StructuralElementHeader}, JustificationValues.Center, "StructuralElementHeaderNotCentered"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Appendix, OfNumbering: null}, 709, 0, "IncorrectBodyTextFirstLineIndent", "IncorrectBodyTextLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 0}, 709, 0, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 1}, -709, 1418, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 2}, -851, 1560, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphLineSpacingLint(
            // TODO: enforce this for numberings.
            // TODO: enforce this for table cell content, headers, captions.
            x => x is {Class: ParagraphClass.BodyText, OfNumbering: null, OfStructuralElement: not StructuralElement.Appendix},
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
                    x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not (StructuralElement.Bibliography or StructuralElement.Appendix), OfNumbering: null},
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
        new IncorrectFigureNumberingLint(),
        new FigureNotReferencedLint(),
        new BibliographySourceNotReferencedLint(),
        new IncorrectHeadingTextLint(),
        new NotEnoughSourcesLint(40, "NotEnoughSources", "NoBibliography"),
        new TextFontLint(),
        new FontSizeLint(x => x is {Class: ParagraphClass.Heading or ParagraphClass.BodyText}, 24, "IncorrectFontSize"),
        new ForceBoldLint(true, x => x is { Class: ParagraphClass.Heading, OutlineLevel: null or < 2 } or {Class: ParagraphClass.StructuralElementHeader}, "HeadingNotBold"),
        new ForceBoldLint(false, x => x is { OutlineLevel: >= 2 }, "SubSubHeadingBold"),
        new ForceBoldLint(false, x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Bibliography}, "BodyTextBold"),
    ];

    public void Run(LintContext ctx)
    {
        foreach (var lint in _lints)
        {
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
}