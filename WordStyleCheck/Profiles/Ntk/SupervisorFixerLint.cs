using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class SupervisorFixerLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectSupervisorFormatting"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!tool.GetFeature(NtkParagraphData.Key)!.IsSupervisorData) continue;
            
            RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

            int start = rat.Text.IndexOf(':') + 1;
            
            while (char.IsWhiteSpace(rat.Text[start])) start++;

            string name = rat.Text.Substring(start);
            string original = name;

            name = name
                .Replace("ассистент", "асс.")
                .Replace("старший преподаватель", "ст.преп.")
                .Replace("доцент", "доц.")
                .Replace("профессор", "проф.")
                
                .Replace("старший научный сотрудник", "с.н.с.")
                .Replace("научный сотрудник", "н.с.")
                
                .Replace("научный сотрудник", "н.с.");

            void OfSciences(string science, string s)
            {
                name = name.Replace($"доктор {science} наук", $"д.{s}.н.")
                    .Replace($"кандидат {science} наук", $"к.{s}.н.");
            }
            
            OfSciences("военных", "в");
            OfSciences("технический", "т");
            OfSciences("физико-математических", "ф.-м");
            OfSciences("экономических", "э");
            OfSciences("химических", "х");
            OfSciences("исторических", "и");
            OfSciences("юридических", "ю");
            OfSciences("социологических", "с");
            OfSciences("фармацевтических", "фарм");

            if (original != name)
            {
                var span = rat.GetSpan(start, rat.Text.Length - start);

                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("IncorrectSupervisorFormatting", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = name,
                            ["Actual"] = original
                        }
                    });
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    span.Replace(name);
                }
            }
        }
    }
}