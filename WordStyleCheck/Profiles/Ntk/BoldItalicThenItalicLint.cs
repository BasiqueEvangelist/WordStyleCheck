using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class BoldItalicThenItalicLint(Predicate<ParagraphPropertiesTool> predicate, string header, string headerMessageId, string bodyMessageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [headerMessageId, bodyMessageId];
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;

            RunAssociatedText assoc = RunAssociatedText.FromParagraph(tool);
            
            if (!assoc.Text.StartsWith(header)) continue; // ???

            var headerSpan = assoc.GetSpan(0, header.Length);
            var bodySpan = assoc.GetSpan(header.Length, assoc.Text.Length - header.Length);
            
            if (!headerSpan.Matches(x => x.Bold && x.Italic))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(headerMessageId, DiagnosticType.FormattingError, new RunSpanDiagnosticContext(headerSpan)));
                }
                else
                {
                    ctx.MarkAutoFixed();

                    foreach (var run in headerSpan.Isolate())
                    {
                        run.RunProperties ??= new RunProperties();

                        run.RunProperties.Italic = new Italic { Val = true };
                        run.RunProperties.Bold = new Bold { Val = true };
                    }
                }
            }
            
            if (!bodySpan.Matches(x => x.Italic && !x.Bold))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(bodyMessageId, DiagnosticType.FormattingError,  new RunSpanDiagnosticContext(bodySpan)));
                }
                else
                {
                    ctx.MarkAutoFixed();

                    foreach (var run in bodySpan.Isolate())
                    {
                        run.RunProperties ??= new RunProperties();

                        run.RunProperties.Italic = new Italic { Val = true };
                        run.RunProperties.Bold = new Bold { Val = false };
                    }
                }
            }
        }
    }
}