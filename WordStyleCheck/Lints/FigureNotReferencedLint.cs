using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureNotReferencedLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData?.Type != CaptionType.Figure) continue;

            var refOptions = tool.CaptionData.Value.GetReferenceTexts();
            
            if (refOptions.Count == 0) continue;

            bool wasReferenced = false;
            foreach (var other in ctx.Document.AllParagraphs)
            {
                if (other == p) continue;
                
                // TODO: SearchValues?

                var text = Utils.CollectParagraphText(other);
                
                foreach (var option in refOptions)
                {
                    if (text.Contains(option, StringComparison.InvariantCultureIgnoreCase))
                    {
                        wasReferenced = true;
                        break;
                    }
                }

                if (wasReferenced) break;
            }
            
            if (wasReferenced) continue;
            
            ctx.AddMessage(new LintMessage("Figure was never referenced in text", new ParagraphDiagnosticContext(p)));
        }
    }
}