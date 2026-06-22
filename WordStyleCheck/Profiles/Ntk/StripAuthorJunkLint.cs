using System.Text.RegularExpressions;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class StripAuthorJunkLint : ILint
{
    private static readonly Regex NameRegex = new Regex(NtkPartsClassifier.NameRegexText);
    
    private static readonly Regex NameFullRegex = new Regex(@"([А-Я][а-я]+)\s+([А-Я][а-я]+)\s*([А-Я][а-я]+)");
    
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["AuthorJunk"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (!tool.GetFeature(NtkParagraphData.Key)!.IsAuthorData) continue;

            List<string> names = [];

            string trimmed = tool.Contents.Trim();
            
            foreach (var match in NameRegex.EnumerateMatches(trimmed))
            {
                names.Add(trimmed.Substring(match.Index, match.Length));
            }

            foreach (var match in NameFullRegex.Matches(trimmed).Cast<Match>())
            {
                string correctName = match.Groups[1].Value + " " + match.Groups[2].Value[0] + "." + " " +
                                 match.Groups[3].Value[0] + ".";
                
                names.Add(correctName);
            }

            string correct = string.Join(", ", names);
            
            if (trimmed == correct) continue;
            
            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic("AuthorJunk", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = correct,
                        ["Actual"] = trimmed
                    }
                });
            }
            else
            {
                ctx.MarkAutoFixed();
                    
                RunAssociatedText.FromParagraph(tool).GetSpan(0, tool.Contents.Length).Replace(correct);
            }
        }
    }
}