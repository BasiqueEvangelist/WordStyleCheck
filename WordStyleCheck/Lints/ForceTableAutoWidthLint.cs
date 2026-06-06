using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceTableAutoWidthLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["TableMustBeAutoWidth"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var table in ctx.Document.AllTables)
        {
            var tool = ctx.Document.GetTool(table);
            
            if (tool.Class is not (TableClass.Table or TableClass.TableContinuation or TableClass.DisplayEquation or TableClass.Listing or TableClass.Figure)) continue;

            if (tool.WidthType != TableWidthUnitValues.Auto)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("TableMustBeAutoWidth", DiagnosticType.FormattingError, new TableDiagnosticContext(table)));
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    // TODO: generate-revisions
                    table.TableProperties ??= new TableProperties();
                    table.TableProperties.TableWidth ??= new TableWidth();
                    table.TableProperties.TableWidth.Type = TableWidthUnitValues.Auto;
                    table.TableProperties.TableWidth.Width = "0";
                }
            }
        }
    }
}