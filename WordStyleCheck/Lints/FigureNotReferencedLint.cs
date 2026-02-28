using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureNotReferencedLint : ILint
{
    //        return [$"(рис. {Number})", $"рисунок {Number}", $"рисунке {Number}", $"рисунком {Number}"];
    private static readonly Regex[] FigureRegexes =
    [
        new("\\(рис\\. ([0-9\\.]+(?:, [0-9\\.]+)*)\\)", RegexOptions.Compiled),
        new("рисунок ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунки ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунке ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунком ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
        new("рисунках ([0-9\\.]+(?:, [0-9\\.]+)*)(?:и ([0-9\\.]+))?", RegexOptions.Compiled),
    ];
    
    public void Run(LintContext ctx)
    {
        HashSet<string> referencedNumbers = [];
        
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
                        referencedNumbers.Add(res.TrimEnd('.'));
                    }
                    
                    if (match.Groups.Count >= 3)
                        referencedNumbers.Add(match.Groups[2].Value.TrimEnd('.'));
                }
            }
        }
        
        Console.WriteLine("Referenced figures: " + string.Join(", ", referencedNumbers));
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData?.Type != CaptionType.Figure) continue;

            if (referencedNumbers.Contains(tool.CaptionData.Value.Number)) continue;
            
            ctx.AddMessage(new LintMessage("FigureNotReferenced", new ParagraphDiagnosticContext(p)));
        }
    }
}