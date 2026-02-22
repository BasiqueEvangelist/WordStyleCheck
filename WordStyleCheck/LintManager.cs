using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager
{
    private readonly List<ILint> _lints =
    [
        new NeedlessParagraphLint(),
        new ParagraphFirstLineIndentLint(),
        new ParagraphSpacingLint(),
        new BodyTextFontLint()
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