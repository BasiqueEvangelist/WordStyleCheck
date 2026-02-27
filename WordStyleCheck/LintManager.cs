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
        new HomemadeListLint(),
        new NeedlessParagraphLint(),
        new ForceJustificationLint(x => x is {Class: ParagraphClass.StructuralElementHeader}, JustificationValues.Center, "StructuralElementHeaderNotCentered"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Appendix, OfNumbering: null}, 709, 0, "IncorrectBodyTextFirstLineIndent", "IncorrectBodyTextLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 0}, 709, 0, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 1}, -709, 1418, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x is {Class: ParagraphClass.Heading, OutlineLevel: 2}, -851, 1560, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphLineSpacingLint(),
        new CorrectStructuralElementHeaderLint(),
        new WrongCaptionPositionLint(CaptionType.Table, false, "IncorrectTableCaptionPosition"),
        new WrongCaptionPositionLint(CaptionType.Figure, true, "IncorrectFigureCaptionPosition"),
        new IncorrectCaptionTextLint(),
        new IncorrectFigureNumberingLint(),
        new FigureNotReferencedLint(),
        new BibliographySourceNotReferencedLint(),
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
                lint.Run(ctx);
            }
        }

        using (new LoudStopwatch("RunLintMerger.Run")) 
            RunLintMerger.Run(ctx.Messages);
        
        using (new LoudStopwatch("ParagraphLintMerger.Run")) 
            ParagraphLintMerger.Run(ctx.Messages);
    }
}