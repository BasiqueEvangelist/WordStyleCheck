using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HandmadeListLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["HandmadeList"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var list in ctx.Document.HandmadeLists)
        {
            ctx.AddMessage(new LintDiagnostic("HandmadeList", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(list.Paragraphs)));
        }
    }
    
}