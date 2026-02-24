using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager
{
    private readonly List<ILint> _lints =
    [
        new PageSizeLint(),
        new NeedlessParagraphLint(),
        new ParagraphFirstLineIndentLint(),
        new ParagraphSpacingLint(),
        new CorrectStructuralElementHeaderLint(),
        new WrongCaptionPositionLint(CaptionType.Table, false, "Table captions must be above their respective tables"),
        new WrongCaptionPositionLint(CaptionType.Figure, true, "Figure captions must be below their respective figures"),
        new BodyTextFontLint(),
        new ForceBoldLint(true, x => x is { Class: ParagraphPropertiesTool.ParagraphClass.Heading, OutlineLevel: null or < 2 }, "Heading text must be bold"),
        new ForceBoldLint(false, x => x is {Class: ParagraphPropertiesTool.ParagraphClass.BodyText, OfStructuralElement: not StructuralElement.Bibliography}, "Body text must not be bold"),
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