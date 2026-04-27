using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceCapsLint(Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
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
            
            if (pTool.Contents.ToUpperInvariant() == pTool.Contents) continue;

            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError,
                    new ParagraphDiagnosticContext(p)));
            }
            else
            {
                ctx.MarkAutoFixed();

                // TODO: generate revisions.
                
                foreach (var run in Utils.DirectRunChildren(p))
                {
                    foreach (var text in run.ChildElements.OfType<Text>())
                    {
                        text.Text = text.Text.ToUpperInvariant();
                    }
                }
            }
        }
    }
}