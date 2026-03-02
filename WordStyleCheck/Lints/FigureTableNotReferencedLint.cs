using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureTableNotReferencedLint : ILint
{
    private static readonly Regex[] FigureRegexes =
    [
        new("\\(рис\\. ([0-9\\.]+(?:, [0-9\\.]+)*)\\)", RegexOptions.Compiled),
        new("рисунок ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунки ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунке ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунком ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунках ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
    ];

    private static readonly Regex[] TableRegexes =
    [
        new("таблица ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("таблицы ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("таблице ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
    ];

    public void Run(LintContext ctx)
    {
        HashSet<string> referencedFigureNumbers = [];
        HashSet<string> referencedTableNumbers = [];

        foreach (var other in ctx.Document.AllParagraphs)
        {
            if (ctx.Document.GetTool(other).Class == ParagraphClass.Caption) continue;
                
            var text = Utils.CollectParagraphText(other);
                
            foreach (var option in FigureRegexes)
            {
                foreach (Match match in option.Matches(text))
                {
                    foreach (var res in match.Groups[1].Value.Split(", "))
                    {
                        referencedFigureNumbers.Add(res.TrimEnd('.'));
                    }
                    
                    if (match.Groups.Count >= 3)
                        referencedFigureNumbers.Add(match.Groups[2].Value.TrimEnd('.'));
                }
            }

            foreach (var option in TableRegexes)
            {
                foreach (Match match in option.Matches(text))
                {
                    foreach (var res in match.Groups[1].Value.Split(", "))
                    {
                        referencedTableNumbers.Add(res.TrimEnd('.'));
                    }

                    if (match.Groups.Count >= 3)
                        referencedTableNumbers.Add(match.Groups[2].Value.TrimEnd('.'));
                }
            }
        }
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if ((tool.CaptionData?.Type) == CaptionType.Figure)
            {
                if (referencedFigureNumbers.Contains(tool.CaptionData.Value.Number)) continue;

                ctx.AddMessage(new LintMessage("FigureNotReferenced", new ParagraphDiagnosticContext(p)));
            }
            else if (tool is { CaptionData: { Type: CaptionType.Table, IsContinuation: false } })
            {
                if (referencedTableNumbers.Contains(tool.CaptionData.Value.Number)) continue;

                ctx.AddMessage(new LintMessage("TableNotReferenced", new ParagraphDiagnosticContext(p)));
            }
        }
    }
}