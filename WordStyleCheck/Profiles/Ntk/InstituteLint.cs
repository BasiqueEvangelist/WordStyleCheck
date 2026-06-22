using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class InstituteLint : ILint
{
    private static readonly string[][] InstituteNames = [
        ["Военный учебный центр при РТУ МИРЭА"],
        ["Институт искусственного интеллекта", "ИИИ"],
        ["Институт информационных технологий", "ИИТ"],
        ["Институт кибербезопасности и цифровых технологий", "ИКБ"],
        ["Институт перспективных технологий и индустриального программирования", "ИПТИП"],
        ["Институт технологий управления", "ИТУ"],
        ["Институт тонких химических технологий имени \u00a0М.В.\u00a0Ломоносова", "ИТХТ"],
        ["Институт радиоэлектроники и информатики", "ИРИ"],
        ["Передовая инженерная школа СВЧ-электроники"]
    ];
    
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectInstitute", "IncorrectInstituteFixed"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!tool.GetFeature(NtkParagraphData.Key)!.IsSourceInstitute) continue;

            RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

            int startIndex = rat.Text.IndexOf("РТУ МИРЭА", StringComparison.InvariantCultureIgnoreCase) + "РТУ МИРЭА".Length;
            
            while (!char.IsLetter(rat.Text[startIndex])) startIndex += 1;
            
            var span = rat.GetSpan(startIndex, rat.Text.Length - startIndex);
            
            string text = span.ToString();
            
            if (InstituteNames.Select(x => x[0]).Any(x => x == text)) continue;
            
            var candidates = InstituteNames
                .Select(x => new { Correct = x[0], Distance = x.Select(y => StringEx.LevenshteinDistance(text, y)).Min() })
                .Where(x => x.Distance <= text.Length / 5)
                .ToList();
                
            string? best = candidates
                .MinBy(x => x.Distance)
                ?.Correct;

            if (best == null)
            {
                ctx.AddMessage(new LintDiagnostic("IncorrectInstitute", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span))
                {
                    Parameters = new()
                    {
                        ["Actual"] = text
                    }
                });
            }
            else
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("IncorrectInstituteFixed", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = best,
                            ["Actual"] = text
                        }
                    });
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    rat.GetSpan(0, rat.Text.Length).Replace("РТУ МИРЭА, " + best);
                }
            }
        }
    }
}