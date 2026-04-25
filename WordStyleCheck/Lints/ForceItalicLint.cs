using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceItalicLint(bool italic, Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
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
            
            foreach (var r in Utils.DirectRunChildren(p))
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.Italic != italic)
                {
                    ctx.AddMessage(new LintMessage(messageId, new RunDiagnosticContext(r))
                    {
                        AutoFix = () =>
                        {
                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            
                            if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);

                            if (r.RunProperties.Italic == null) r.RunProperties.Italic = new Italic();
                            
                            r.RunProperties.Italic.Val = italic;
                        }
                    });
                }
            }
        }
    }
}