using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager
{
    private readonly List<ILint> _lints =
    [
        new NeedlessParagraphLint(),
        new ParagraphFirstLineIndentLint(),
        new ParagraphSpacingLint(),
        new BodyTextFontLint(),
        new ForceBoldLint(true, x => x is { Class: ParagraphPropertiesTool.ParagraphClass.Heading, OutlineLevel: null or < 2 }, "Heading text must be bold"),
        // TODO: make this lint not fire for text splitting different 
        new ForceBoldLint(false, x => x.Class == ParagraphPropertiesTool.ParagraphClass.BodyText, "Body text must not be bold"),
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