using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class AbstractStartsWithLowercaseLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["AbstractStartsWithUppercase"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (!tool.GetFeature(NtkParagraphData.Key)!.IsAbstract) continue;

            if (!tool.Contents.StartsWith("Аннотация: ")) continue;

            var rat = RunAssociatedText.FromParagraph(tool);

            int start = rat.Text.IndexOf(':') + 1;

            while (start < rat.Text.Length && char.IsWhiteSpace(rat.Text[start])) start++;
            
            if (start >= rat.Text.Length) continue;

            if (!char.IsUpper(rat.Text[start])) continue;

            var span = rat.GetSpan(start, 1);
            
            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic("AbstractStartsWithUppercase", DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span)));
            }
            else
            {
                ctx.MarkAutoFixed();
                
                span.Replace(char.ToLower(rat.Text[start]).ToString());
            }
        }
    }
}