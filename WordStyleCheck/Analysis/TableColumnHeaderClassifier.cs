using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class TableColumnHeaderClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        foreach (var p in ctx.AllParagraphs)
        {
            var tool = ctx.GetTool(p);

            if (tool.ContainingTableRow is not { } tr) continue;

            var table = (Table?)tr.Parent;

            if (table == null || ctx.GetTool(table).Class != TableClass.Table) continue;

            int rowIndex = table.ChildElements.ToList().IndexOf(tr);

            if (rowIndex == 0)
            {
                tool.ProbablyTableColumnHeader = true;
            }
        }
    }
}