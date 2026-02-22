using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceBoldLint(bool bold, Predicate<ParagraphPropertiesTool> predicate, string message) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);

            if (!predicate(pTool))
            {
                continue;
            }
            
            foreach (var r in p.Descendants<Run>())
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.Bold != bold)
                {
                    ctx.AddMessage(new LintMessage(message, new RunDiagnosticContext(r))
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