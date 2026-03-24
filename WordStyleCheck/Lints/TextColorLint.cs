using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class TextColorLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["TextNotAutoColor", "TextHighlighted"];

    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);

            if (pTool.IsEmptyOrDrawing) continue;

            if (pTool.Class == ParagraphClass.CodeListing) continue;

            if (pTool.Class == ParagraphClass.InsideDrawing) continue;

            foreach (var r in Utils.DirectRunChildren(p))
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;

                RunPropertiesTool tool = ctx.Document.GetTool(r);
                
                if (tool.IsHyperlink) continue;

                if ((tool.Color ?? "auto") is not ("auto" or "000000"))
                {
                    ctx.AddMessage(new LintMessage("TextNotAutoColor", new RunDiagnosticContext(r))
                    {
                        AutoFix = () =>
                        {
                            // TODO: generate revisions

                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            r.RunProperties.Color = new Color()
                            {
                                Val = "auto"
                            };
                        }
                    });
                }

                if (tool.Highlight != HighlightColorValues.None)
                {
                    ctx.AddMessage(new LintMessage("TextHighlighted", new RunDiagnosticContext(r))
                    {
                        AutoFix = () =>
                        {
                            // TODO: generate revisions

                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            r.RunProperties.Highlight = null;
                        }
                    });
                }
            }
        }
    }
}
