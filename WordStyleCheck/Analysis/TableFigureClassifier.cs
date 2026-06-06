using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class TableFigureClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        foreach (var table in ctx.AllTables)
        {
            var tool = ctx.GetTool(table);
            
            if (tool.Class != TableClass.Unknown) continue;
            if (tool.Rows.Count != 2 || tool.ColumnCount != 1) continue;
            if (!CaptionClassifierData.HasFigure(tool.Rows[0].Row)) continue;

            var possiblyCaption = ctx.GetTool(tool.Rows[1][0].Descendants<Paragraph>().First());
            
            if (CaptionClassifierData.TryParse(possiblyCaption, CaptionType.Figure, true, null) is {} data)
            {
                possiblyCaption.CaptionData = data;
                tool.Caption = possiblyCaption;
            }
        }
    }
}