using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceCapsLint(Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(LintContext ctx)
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
            
            if (pTool.Contents.ToUpperInvariant() == pTool.Contents) continue;
            
            ctx.AddMessage(new LintDiagnostic(messageId, new ParagraphDiagnosticContext(p), AutoFix: () =>
            {
                foreach (var run in Utils.DirectRunChildren(p))
                {
                    foreach (var text in run.ChildElements.OfType<Text>())
                    {
                        text.Text = text.Text.ToUpperInvariant();
                    }
                }
            }));
        }
    }
}