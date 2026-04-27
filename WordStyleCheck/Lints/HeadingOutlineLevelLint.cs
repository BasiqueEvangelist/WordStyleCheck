using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HeadingOutlineLevelLint(Predicate<ParagraphPropertiesTool> requiresOutlineLevel) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["NonHeadingWithOutlineLevel", "HeadingWithoutOutlineLevel", "IncorrectHeadingOutlineLevel"];

    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (tool.IsIgnored) continue;

            if (tool is { OutlineLevel: not null } && !requiresOutlineLevel(tool))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("NonHeadingWithOutlineLevel", DiagnosticType.FormattingError,
                        new ParagraphDiagnosticContext(p)));
                }
                else
                {
                    ctx.MarkAutoFixed();

                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.OutlineLevel ??= new OutlineLevel();
                    p.ParagraphProperties.OutlineLevel.Val = 9;
                }
            }

            if (tool is { HeadingData: not null, OutlineLevel: null })
            {
                if (tool.HeadingData.Level >= 4) continue;
                
                ctx.AddMessage(new LintDiagnostic("HeadingWithoutOutlineLevel", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
            }

            if (tool is { HeadingData.Level: var level, HeadingData.IsConclusion: false, OutlineLevel: { } outlineLevel } && outlineLevel + 1 != level)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("IncorrectHeadingOutlineLevel", DiagnosticType.FormattingError,
                        new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = level.ToString(),
                            ["Actual"] = (outlineLevel + 1).ToString()
                        }
                    });
                }
                else
                {
                    ctx.MarkAutoFixed();

                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.OutlineLevel ??= new OutlineLevel();
                    p.ParagraphProperties.OutlineLevel.Val = level - 1;
                }
            }
        }
    }
}
