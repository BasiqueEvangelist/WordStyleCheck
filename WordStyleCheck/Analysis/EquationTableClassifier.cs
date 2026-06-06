using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeMathParagraph = DocumentFormat.OpenXml.Math.Paragraph;


namespace WordStyleCheck.Analysis;

public class EquationTableClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        foreach (var table in ctx.AllTables)
        {
            var tool = ctx.GetTool(table);
            
            if (tool.Class != TableClass.Unknown) continue;
            if (tool.Rows.Count != 1 || tool.ColumnCount is not (1 or 2)) continue;

            var eqParagraph = tool.Rows[0][0].Descendants<Paragraph>().First();

            OpenXmlElement? oMath = null;
            
            foreach (var c1 in eqParagraph.ChildElements)
            {
                if (c1 is Run r)
                {
                    foreach (var c2 in r.ChildElements)
                    {
                        if (c2 is Text t && !string.IsNullOrWhiteSpace(t.Text))
                            goto outer;
                    }
                }

                if (c1 is DocumentFormat.OpenXml.Math.OfficeMath or OfficeMathParagraph)
                {
                    oMath = c1;
                    break;
                }

                if (oMath != null) break;
            }
            
            if (oMath == null) continue;

            string? number = null;
            
            if (tool.ColumnCount == 2)
            {
                var contents = ctx.GetTool(tool[0][1].Descendants<Paragraph>().First()).Contents;
                
                if (contents.Length > 0 && contents[0] == '(' && contents[^1] == ')')
                {
                    number = contents.Substring(1, contents.Length - 2);
                }
            }

            tool.EquationData = new EquationClassifierData()
            {
                MathElement = oMath,
                Number = number
            };
            
            outer: ;
        }
    }
}