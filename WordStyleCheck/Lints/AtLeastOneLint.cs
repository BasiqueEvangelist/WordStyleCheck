using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class AtLeastOneLint(Predicate<ParagraphPropertiesTool> predicate, string messageId, DiagnosticType type, bool atEnd) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];

    public void Run(LintContext ctx)
    {
        bool any = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Any(predicate.Invoke);

        if (any) return;
        
        ctx.AddMessage(new LintDiagnostic(
            messageId,
            type,
            atEnd 
                ? new EndOfDocumentDiagnosticContext(ctx.Document.AllParagraphs.Last())
                : new StartOfDocumentDiagnosticContext(ctx.Document.AllParagraphs.First())));
    }
}