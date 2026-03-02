using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureTableNotReferencedLint : ILint
{
    [StringSyntax("Regex")]
    private const string Reference = @"(?:[А-Я]\.)?[0-9]+(?:\.[0-9]+)*";
    
    [StringSyntax("Regex")]
    private const string ReferenceSpan = $@"({Reference})\s*(?:-|–)\s*({Reference})";

    private static readonly Regex ReferenceRegex =
        new(Reference, RegexOptions.Compiled | RegexOptions.IgnoreCase);
    
    private static readonly Regex ReferenceSpanRegex =
        new(ReferenceSpan, RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private const string ReferenceOrSpan = $@"{Reference}(\s*(?:-|–)\s*({Reference}))?";
    
    [StringSyntax("Regex")]
    private const string ReferenceList = $@"{ReferenceOrSpan}(?:,\s+{ReferenceOrSpan})*(?:и\s+{ReferenceOrSpan})?"; 
    
    private static readonly Regex[] FigureRegexes =
    [
        new($"рис\\.\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"рисунок\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"рисунки\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"рисунке\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"рисунком\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"рисунках\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
    ];

    private static readonly Regex[] TableRegexes =
    [
        new($"табл\\.\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"таблица\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"таблицы\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
        new($"таблице\\s+{ReferenceList}", RegexOptions.Compiled | RegexOptions.IgnoreCase),
    ];

    public void Run(LintContext ctx)
    {
        Dictionary<string, Paragraph> referencedFigureNumbers = [];
        Dictionary<string, Paragraph> referencedTableNumbers = [];

        foreach (var other in ctx.Document.AllParagraphs)
        {
            if (ctx.Document.GetTool(other).Class == ParagraphClass.Caption) continue;

            var text = Utils.CollectParagraphText(other);

            void AddMatches(Dictionary<string, Paragraph> referencedNumbers, Regex[] regexes)
            {
                foreach (var option in regexes)
                {
                    foreach (Match match in option.Matches(text))
                    {
                        string matched = match.Value;

                        foreach (Match subMatch in ReferenceRegex.Matches(matched))
                        {
                            string referenced = subMatch.Value.TrimEnd('.');

                            referencedNumbers.TryAdd(referenced, other);
                        }

                        foreach (Match subMatch in ReferenceSpanRegex.Matches(matched))
                        {
                            string start = subMatch.Groups[1].Value;
                            string end = subMatch.Groups[2].Value;

                            string[] startSplit = start.Split(".");
                            string[] endSplit = end.Split(".");

                            if (startSplit.Length != endSplit.Length) continue;
                            if (!startSplit[..^1].SequenceEqual(endSplit[..^1])) continue;

                            if (!int.TryParse(startSplit[^1], out var startNum)) continue;
                            if (!int.TryParse(endSplit[^1], out var endNum)) continue;

                            for (int i = startNum; i <= endNum; i++)
                            {
                                startSplit[^1] = i.ToString();

                                string referenced = string.Join(".", startSplit);

                                referencedNumbers.TryAdd(referenced, other);
                            }
                        }
                    }
                }
            }

            AddMatches(referencedFigureNumbers, FigureRegexes);

            AddMatches(referencedTableNumbers, TableRegexes);
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