using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FontSizeLint(Predicate<ParagraphPropertiesTool> predicate, int fontSize, bool force, string messageId) : ILint
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

                if (tool.FontSize != null && (force ? tool.FontSize != fontSize : tool.FontSize < fontSize))
                {
                    ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError, new RunDiagnosticContext(r))
                    {
                        Parameters = new()
                        {
                            ["ExpectedPt"] = (fontSize / 2).ToString(CultureInfo.InvariantCulture),
                            ["ActualPt"] = (tool.FontSize.Value / 2).ToString(CultureInfo.InvariantCulture)
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