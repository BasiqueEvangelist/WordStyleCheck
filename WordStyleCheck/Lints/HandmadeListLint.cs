using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HandmadeListLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var list in ctx.Document.HandmadeLists)
        {
            ctx.AddMessage(new LintMessage("HandmadeList", new ParagraphDiagnosticContext(list.Paragraphs)));
        }
    }
    
}