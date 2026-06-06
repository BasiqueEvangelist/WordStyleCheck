using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class InstituteCapitalizationLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectInstituteCapitalization"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!tool.GetFeature(NtkParagraphData.Key)!.IsSourceInstitute) continue;

            RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

            int startIndex = rat.Text.IndexOf("Институт", StringComparison.InvariantCultureIgnoreCase);
            var span = rat.GetSpan(startIndex, rat.Text.Length - startIndex);
            string text = span.ToString();
            string correct = char.ToUpper(text[0]) + text[1..].ToLower();

            if (text != correct)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("IncorrectInstituteCapitalization", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span)));
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    span.Replace(correct);
                }
            }
        }
    }
}