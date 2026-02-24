using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class WrongCaptionPositionLint(CaptionType captionType, bool shouldBeBelow, string message) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData?.Type != captionType) continue;
            
            if (tool.CaptionData.Value.IsBelow == shouldBeBelow) continue;
            
            ctx.AddMessage(new LintMessage(message, new ParagraphDiagnosticContext(p))
            {
                AutoFix = () =>
                {
                    p.Remove();

                    if (shouldBeBelow)
                    {
                        tool.CaptionData.Value.TargetedElement.InsertAfterSelf(p);
                    }
                    else
                    {
                        tool.CaptionData.Value.TargetedElement.InsertBeforeSelf(p);
                    }
                }
            });
        }
    }
}