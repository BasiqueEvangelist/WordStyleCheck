using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForbidLint(Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;
            
            ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.ContentError, new ParagraphDiagnosticContext(p)));

            if (ctx.AutomaticallyFix)
            {
                Utils.HighlightRed(p, ctx.GenerateRevisions);
            }
        }
    }
}