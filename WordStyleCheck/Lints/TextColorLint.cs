using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class TextColorLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["TextNotAutoColor", "TextHighlighted"];

    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);

            if (pTool.IsEmptyOrDrawing) continue;

            if (pTool.ProbablyCodeListing) continue;

            if (pTool.ContainingTextBox != null) continue;

            foreach (var r in Utils.DirectRunChildren(p))
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;

                RunPropertiesTool tool = ctx.Document.GetTool(r);
                
                if (tool.IsHyperlink) continue;

                if ((tool.Color ?? "auto") is not ("auto" or "000000"))
                {
                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic("TextNotAutoColor", DiagnosticType.FormattingError,
                            new RunDiagnosticContext(r)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();

                        r.RunProperties ??= new RunProperties();
                        if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);
                        
                        r.RunProperties.Color = new Color()
                        {
                            Val = "auto"
                        };
                    }
                }

                if (tool.Highlight != HighlightColorValues.None)
                {
                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic("TextHighlighted", DiagnosticType.FormattingError,
                            new RunDiagnosticContext(r)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();

                        r.RunProperties ??= new RunProperties();
                        Utils.SnapshotRunProperties(r.RunProperties);
                        
                        r.RunProperties.Highlight = null;
                    }
                }
            }
        }
    }
}
