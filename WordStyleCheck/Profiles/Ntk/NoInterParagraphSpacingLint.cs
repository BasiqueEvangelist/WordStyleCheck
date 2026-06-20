using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class NoInterParagraphSpacingLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["ForbiddenInterParagraphSpacing"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.GetFeature(NtkParagraphData.Key)!.IsProbablyJunk) continue;

            int after = tool.ActualAfterSpacing ?? 0;
            int before = tool.ActualBeforeSpacing ?? 0;

            if (after != 0 || before != 0)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("ForbiddenInterParagraphSpacing", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
                    p.ParagraphProperties.SpacingBetweenLines.After = "0";
                    p.ParagraphProperties.SpacingBetweenLines.AfterLines = null;
                    p.ParagraphProperties.SpacingBetweenLines.AfterAutoSpacing = null;
                    p.ParagraphProperties.SpacingBetweenLines.Before = "0";
                    p.ParagraphProperties.SpacingBetweenLines.BeforeLines = null;
                    p.ParagraphProperties.SpacingBetweenLines.BeforeAutoSpacing = null;
                }
            }
        }
    }
}