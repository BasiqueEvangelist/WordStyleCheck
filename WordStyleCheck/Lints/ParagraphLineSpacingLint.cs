using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphLineSpacingLint(Predicate<ParagraphPropertiesTool> predicate, int lineSpacing, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];

    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;

            if (!predicate(tool)) continue;
            
            if (tool.LineSpacing == null) continue;
            
            if (tool.LineSpacing != lineSpacing)
            {
                ctx.AddMessage(new LintDiagnostic(
                    messageId,
                    DiagnosticType.FormattingError,
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = (lineSpacing / 240.0).ToString(CultureInfo.InvariantCulture),
                            ["Actual"] = (tool.LineSpacing.Value / 240.0).ToString(CultureInfo.InvariantCulture)
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