using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FontSizeLint(Predicate<ParagraphPropertiesTool> predicate, int fontSize, bool force, string messageId)
    : ParagraphLint
{
    public override IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];

    public override void Run(ILintContext ctx, ParagraphPropertiesTool p)
    {
        if (!predicate(p))
        {
            return;
        }
        
        foreach (var r in Utils.DirectRunChildren(p.Paragraph))
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;

            RunPropertiesTool tool = ctx.Document.GetTool(r);

            if (tool.FontSize != null && (force ? tool.FontSize != fontSize : tool.FontSize < fontSize))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError,
                        new RunDiagnosticContext(r))
                    {
                        Parameters = new()
                        {
                            ["ExpectedPt"] = (fontSize / 2).ToString(CultureInfo.InvariantCulture),
                            ["ActualPt"] = (tool.FontSize.Value / 2).ToString(CultureInfo.InvariantCulture)
                        }
                    });
                }
                else
                {
                    ctx.MarkAutoFixed();

                    r.RunProperties ??= new RunProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);

                    r.RunProperties.FontSize ??= new FontSize();
                    r.RunProperties.FontSize.Val = fontSize.ToString();
                }
            }
        }
    }
}