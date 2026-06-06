using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ListingFormattingLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["ListingMustBeOnceCell", "ListingMustBeMonospace"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var table in ctx.Document.AllTables)
        {
            var tool = ctx.Document.GetTool(table);
            
            if (tool.Class != TableClass.Listing) continue;

            if (tool.Rows.Count != 1 || tool.ColumnCount is not (1 or 2))
            {
                ctx.AddMessage(new LintDiagnostic("ListingMustBeOneCell", DiagnosticType.ContentError, new TableDiagnosticContext(table)));
            }

            foreach (var r in table.Descendants<Run>())
            {
                var rTool = ctx.Document.GetTool(r);

                if (Utils.IsMonospaceFont(rTool.AsciiFont ?? "<?>")) continue;
                
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("ListingMustBeMonospace", DiagnosticType.FormattingError,
                        new RunDiagnosticContext(r)));
                }
                else
                {
                    ctx.MarkAutoFixed();
                        
                    r.RunProperties ??= new RunProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(r.RunProperties);
                        
                    r.RunProperties.RunFonts ??= new RunFonts();
                    r.RunProperties.RunFonts.Ascii = "Courier New";
                    r.RunProperties.RunFonts.HighAnsi = "Courier New";
                    r.RunProperties.RunFonts.ComplexScript = "Courier New";
                }
            }
        }
    }
}