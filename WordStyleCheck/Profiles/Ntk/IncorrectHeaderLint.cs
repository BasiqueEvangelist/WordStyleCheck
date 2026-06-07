using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class IncorrectHeaderLint(Predicate<ParagraphPropertiesTool> predicate, List<string> options, string correct,
    string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (!predicate(tool)) continue;

            if (tool.Contents.StartsWith(correct)) continue;

            foreach (var option in options)
            {
                if (tool.Contents.StartsWith(option))
                {
                    RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

                    var span = rat.GetSpan(0, option.Length);

                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError, new RunSpanDiagnosticContext(span)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();
                        span.Replace(correct);
                        
                        tool.ReloadContents();
                    }
                    
                    break;
                }
            }
            
            ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
        }
    }
}