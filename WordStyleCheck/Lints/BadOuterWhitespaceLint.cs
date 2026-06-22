using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class BadOuterWhitespaceLint(Predicate<ParagraphPropertiesTool> predicate) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["BadWhitespaceBeforeText", "BadWhitespaceAfterText"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;
            if (tool.IsIgnored) continue;
            if (tool.IsEmptyOrDrawing) continue;

            var text = RunAssociatedText.FromParagraph(tool);

            {
                int startLength = 0;
                while (startLength < text.Text.Length)
                {
                    if (!char.IsWhiteSpace(text.Text[startLength])) break;

                    startLength += 1;
                }

                if (startLength > 0)
                {
                    var span = text.GetSpan(0, startLength);

                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic("BadWhitespaceBeforeText", DiagnosticType.FormattingError,
                            new RunSpanDiagnosticContext(span)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();

                        foreach (var run in span.Isolate().ToList())
                        {
                            run.Remove();
                        }
                        
                        tool.ReloadContents();
                        text = RunAssociatedText.FromParagraph(tool);
                    }
                }
            }

            {
                int endLength = 0;
                while (endLength < text.Text.Length)
                {
                    if (!char.IsWhiteSpace(text.Text[text.Text.Length - endLength - 1])) break;

                    endLength += 1;
                }

                if (endLength > 0)
                {
                    var span = text.GetSpan(text.Text.Length - endLength, endLength);

                    if (!ctx.AutomaticallyFix)
                    {
                        ctx.AddMessage(new LintDiagnostic("BadWhitespaceAfterText", DiagnosticType.FormattingError,
                            new RunSpanDiagnosticContext(span)));
                    }
                    else
                    {
                        ctx.MarkAutoFixed();

                        foreach (var run in span.Isolate().ToList())
                        {
                            run.Remove();
                        }
                        
                        tool.ReloadContents();
                    }
                }
            }
        }
    }
}