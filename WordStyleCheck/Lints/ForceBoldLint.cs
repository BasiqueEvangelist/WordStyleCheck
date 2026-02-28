using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceBoldLint(bool bold, Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);
            
            if (pTool.IsEmptyOrDrawing)
            {
                continue;
            }

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
                    ctx.AddMessage(new LintMessage(messageId, new RunDiagnosticContext(r))
                    {
                        AutoFix = () =>
                        {
                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            
                            if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);

                            if (r.RunProperties.Bold == null) r.RunProperties.Bold = new Bold();
                            
                            r.RunProperties.Bold.Val = bold;
                        }
                    });
                }
            }
        }
    }
}