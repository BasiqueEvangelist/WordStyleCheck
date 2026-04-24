using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HeadingOutlineLevelLint(Predicate<ParagraphPropertiesTool> requiresOutlineLevel) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["NonHeadingWithOutlineLevel", "HeadingWithoutOutlineLevel", "IncorrectHeadingOutlineLevel"];

    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (tool.IsIgnored) continue;

            if (tool is { OutlineLevel: not null } && !requiresOutlineLevel(tool))
            {
                ctx.AddMessage(new LintMessage("NonHeadingWithOutlineLevel", new ParagraphDiagnosticContext(p))
                {
                    AutoFix = () =>
                    {
                        if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();

                        if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                        p.ParagraphProperties.OutlineLevel?.Remove();

                        p.ParagraphProperties.OutlineLevel = new OutlineLevel()
                        {
                            Val = 9
                        };
                    }
                });
            }

            if (tool is { HeadingData: not null, OutlineLevel: null })
            {
                if (tool.HeadingData.Level >= 4) continue;
                
                ctx.AddMessage(new LintMessage("HeadingWithoutOutlineLevel", new ParagraphDiagnosticContext(p)));
            }

            if (tool is { HeadingData.Level: var level, HeadingData.IsConclusion: false, OutlineLevel: { } outlineLevel } && outlineLevel + 1 != level)
            {
                ctx.AddMessage(new LintMessage("IncorrectHeadingOutlineLevel", new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = level.ToString(),
                        ["Actual"] = (outlineLevel + 1).ToString()
                    },
                    AutoFix = () =>
                    {
                        p.ParagraphProperties ??= new ParagraphProperties();

                        if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                        p.ParagraphProperties.OutlineLevel ??= new OutlineLevel();
                        
                        p.ParagraphProperties.OutlineLevel.Val = level - 1;
                    }
                });
            }
        }
    }
}
