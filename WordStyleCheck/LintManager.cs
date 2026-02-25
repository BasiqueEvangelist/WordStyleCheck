using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager
{
    private readonly List<ILint> _lints =
    [
        new PageSizeLint(),
        new PageMarginsLint(),
        new NeedlessParagraphLint(),
        new ParagraphIndentLint(),
        new ParagraphSpacingLint(),
        new CorrectStructuralElementHeaderLint(),
        new WrongCaptionPositionLint(CaptionType.Table, false, "Table captions must be above their respective tables"),
        new WrongCaptionPositionLint(CaptionType.Figure, true, "Figure captions must be below their respective figures"),
        new IncorrectCaptionTextLint(),
        new FigureNotReferencedLint(),
        new BodyTextFontLint(),
        new FontSizeLint(x => x is {Class: ParagraphClass.Heading or ParagraphClass.BodyText}, 24, "Text font size must be at least 12pt"),
        new ForceBoldLint(true, x => x is { Class: ParagraphClass.Heading, OutlineLevel: null or < 2 }, "Heading text must be bold"),
        new ForceBoldLint(false, x => x is {Class: ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Bibliography}, "Body text must not be bold"),
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