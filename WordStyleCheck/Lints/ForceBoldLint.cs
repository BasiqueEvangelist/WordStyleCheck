using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceBoldLint(bool bold, Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);
            
            if (pTool.IsEmptyOrDrawing) continue;
            if (pTool.IsIgnored) continue;

            if (!predicate(pTool))
            {
                continue;
            }
            
            foreach (var r in Utils.DirectRunChildren(p))
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.Bold != bold)
                {
                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError,
                            new RunDiagnosticContext(r)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();

                        r.RunProperties ??= new RunProperties();
                        if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);

                        r.RunProperties.Bold ??= new Bold();
                        r.RunProperties.Bold.Val = bold;
                    }
                }
            }
        }
    }
}