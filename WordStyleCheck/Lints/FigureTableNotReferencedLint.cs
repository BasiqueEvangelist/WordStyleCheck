using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureTableNotReferencedLint : ILint
{
    private static readonly Regex[] FigureRegexes =
    [
        new("\\(рис\\. ([0-9\\.]+(?:, [0-9\\.]+)*)\\)", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("рисунок ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("рисунки ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("рисунке ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("рисунком ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("рисунках ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
    ];

    private static readonly Regex[] TableRegexes =
    [
        new("табл\\. ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("таблица ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("таблицы ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new("таблице ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled | RegexOptions.IgnoreCase),
    ];

    public void Run(LintContext ctx)
    {
        Dictionary<string, Paragraph> referencedFigureNumbers = [];
        Dictionary<string, Paragraph> referencedTableNumbers = [];

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
                        string referenced = res.TrimEnd('.');

                        if (referencedFigureNumbers.ContainsKey(referenced)) continue;

                        referencedFigureNumbers[referenced] = other;
                    }

                    if (match.Groups.Count >= 3)
                    {
                        string referenced = match.Groups[2].Value.TrimEnd('.');

                        if (referencedFigureNumbers.ContainsKey(referenced)) continue;

                        referencedFigureNumbers[referenced] = other;
                    }
                }
            }

            foreach (var option in TableRegexes)
            {
                foreach (Match match in option.Matches(text))
                {
                    foreach (var res in match.Groups[1].Value.Split(", "))
                    {
                        string referenced = res.TrimEnd('.');

                        if (referencedTableNumbers.ContainsKey(referenced)) continue;

                        referencedTableNumbers[referenced] = other;
                    }

                    if (match.Groups.Count >= 3)
                    {
                        string referenced = match.Groups[2].Value.TrimEnd('.');

                        if (referencedTableNumbers.ContainsKey(referenced)) continue;

                        referencedTableNumbers[referenced] = other;
                    }
                }
            }
        }
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if ((tool.CaptionData?.Type) == CaptionType.Figure)
            {
                if (referencedFigureNumbers.TryGetValue(tool.CaptionData.Value.Number, out var firstMention))
                {
                    var fToplevel = FindTopLevel(p);
                    var mToplevel = FindTopLevel(firstMention);

                    var bodyList = fToplevel.Parent!.ChildElements.ToList();

                    var fIdx = bodyList.IndexOf(fToplevel);
                    var mIdx = bodyList.IndexOf(mToplevel);

                    if (fIdx < mIdx)
                    {
                        ctx.AddMessage(new LintMessage("FigureBeforeFirstReference", new ParagraphDiagnosticContext(p)));
                    }
                }
                else
                {
                    ctx.AddMessage(new LintMessage("FigureNotReferenced", new ParagraphDiagnosticContext(p)));
                }
            }
            else if (tool is { CaptionData: { Type: CaptionType.Table, IsContinuation: false } })
            {
                if (referencedTableNumbers.TryGetValue(tool.CaptionData.Value.Number, out var firstMention))
                {
                    var fToplevel = FindTopLevel(p);
                    var mToplevel = FindTopLevel(firstMention);

                    var bodyList = fToplevel.Parent!.ChildElements.ToList();

                    var fIdx = bodyList.IndexOf(fToplevel);
                    var mIdx = bodyList.IndexOf(mToplevel);

                    if (fIdx < mIdx)
                    {
                        ctx.AddMessage(new LintMessage("TableBeforeFirstReference", new ParagraphDiagnosticContext(p)));
                    }
                }
                else
                {
                    ctx.AddMessage(new LintMessage("TableNotReferenced", new ParagraphDiagnosticContext(p)));
                }
            }
        }
    }

    private static OpenXmlElement FindTopLevel(OpenXmlElement e)
    {
        while (e.Parent is not (null or Body)) e = e.Parent;

        return e;
    }
}