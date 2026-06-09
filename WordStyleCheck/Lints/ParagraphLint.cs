using WordStyleCheck.Analysis;

namespace WordStyleCheck.Lints;

public abstract class ParagraphLint : ILint
{
    public abstract IReadOnlyList<string> EmittedDiagnostics { get; }

    public abstract void Run(ILintContext ctx, ParagraphPropertiesTool p);

    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;

            Run(ctx, tool);
        }
    }
}