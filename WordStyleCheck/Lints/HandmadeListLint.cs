using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HandmadeListLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["HandmadeList"];
    
    public void Run(LintContext ctx)
    {
        foreach (var list in ctx.Document.HandmadeLists)
        {
            ctx.AddMessage(new LintMessage("HandmadeList", new ParagraphDiagnosticContext(list.Paragraphs)));
        }
    }
    
}