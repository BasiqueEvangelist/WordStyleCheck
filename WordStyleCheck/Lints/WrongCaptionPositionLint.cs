using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class WrongCaptionPositionLint(CaptionType captionType, bool shouldBeBelow, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData?.Type != captionType) continue;
            
            if (tool.CaptionData.Value.IsBelow == shouldBeBelow) continue;
            
            if (ctx.AutomaticallyFix && tool.CaptionData.Value.TargetedElement != null)
            {
                ctx.MarkAutoFixed();

                // TODO: generate-revisions
                
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
            else
            {
                ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
            }
        }
    }
}