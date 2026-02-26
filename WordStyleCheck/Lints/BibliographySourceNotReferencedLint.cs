using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class BibliographySourceNotReferencedLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.OfStructuralElement != StructuralElement.Bibliography) continue;
            if (!(tool.OfNumbering is {} numbering)) continue;

            int index = numbering.Paragraphs.IndexOf(p) + 1;
            
            bool wasReferenced = false;
            foreach (var other in ctx.Document.AllParagraphs)
            {
                if (other == p) continue;
                
                // TODO: SearchValues?

                var text = Utils.CollectParagraphText(other);
                
                if (text.Contains("[" + index + "]", StringComparison.InvariantCultureIgnoreCase))
                {
                    wasReferenced = true;
                    break;
                }
            }
            
            if (wasReferenced) continue;
            
            ctx.AddMessage(new LintMessage("BibliographySourceNotReferenced", new ParagraphDiagnosticContext([p], true)));
        }
    }
}