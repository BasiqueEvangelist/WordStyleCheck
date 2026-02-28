using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FontSizeLint(Predicate<ParagraphPropertiesTool> predicate, int fontSize, string messageId) : ILint
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

            foreach (var r in p.Descendants<Run>())
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.FontSize != null && tool.FontSize < fontSize)
                {
                    ctx.AddMessage(new LintMessage(messageId, new RunDiagnosticContext(r))
                    {
                        Parameters = new()
                        {
                            ["ExpectedPt"] = (fontSize / 2).ToString(),
                            ["ActualPt"] = (tool.FontSize.Value / 2).ToString()
                        },
                        AutoFix = () =>
                        {
                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            if (r.RunProperties.FontSize == null) r.RunProperties.FontSize = new FontSize();

                            r.RunProperties.FontSize.Val = fontSize.ToString();
                        }
                    });
                }
            }
        }
    }
}