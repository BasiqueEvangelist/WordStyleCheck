using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceJustificationLint(Predicate<ParagraphPropertiesTool> predicate, JustificationValues justification, string messageId) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;

            if (tool.Justification != justification)
            {
                ctx.AddMessage(new LintMessage(messageId, new ParagraphDiagnosticContext(p))
                {
                    AutoFix = () =>
                    {
                        if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                        if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                        if (p.ParagraphProperties.Justification == null)
                            p.ParagraphProperties.Justification = new Justification();

                        p.ParagraphProperties.Justification.Val = justification;
;                   }
                });
            }
        }
    }
}