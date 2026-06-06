using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class UnknownTableLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["UnknownTable"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var table in ctx.Document.AllTables)
        {
            var tool = ctx.Document.GetTool(table);

            if (tool.Class == TableClass.Unknown)
            {
                ctx.AddMessage(new LintDiagnostic("UnknownTable", DiagnosticType.ContentError,
                    new TableDiagnosticContext(table)));
            }
        }
    }
}