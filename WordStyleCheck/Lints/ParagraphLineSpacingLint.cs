using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphLineSpacingLint(Predicate<ParagraphPropertiesTool> predicate, int lineSpacing, string messageId) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            if (tool.IsEmptyOrDrawing) continue;

            if (!predicate(tool)) continue;
            
            if (tool.LineSpacing == null) continue;
            
            if (tool.LineSpacing != lineSpacing)
            {
                ctx.AddMessage(new LintMessage(
                    messageId,
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = (lineSpacing / 240.0).ToString(CultureInfo.CurrentCulture),
                            ["Actual"] = (tool.LineSpacing.Value / 240.0).ToString(CultureInfo.CurrentCulture)
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                    
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.SpacingBetweenLines == null)
                                p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();

                            p.ParagraphProperties.SpacingBetweenLines.Line = "360";
                        }
                    }
                );
            }
        }
    }
}