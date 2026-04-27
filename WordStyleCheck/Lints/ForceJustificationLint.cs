using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceJustificationLint(Predicate<ParagraphPropertiesTool> predicate, List<JustificationValues> justification, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;
            
            if (tool.IsIgnored || tool.IsEmptyOrDrawing) continue;

            if (!justification.Contains(tool.Justification ?? JustificationValues.Left))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError,
                        new ParagraphDiagnosticContext(p)));

                }
                else
                {
                    ctx.MarkAutoFixed();

                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.Justification ??= new Justification();
                    p.ParagraphProperties.Justification.Val = justification[0];
                }
            }
        }
    }
}