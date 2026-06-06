using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Gost7_32;

public class TextAssociationClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        GostStructuralElement? currentElement = null;
        ParagraphPropertiesTool? currentHeading1 = null;

        foreach (var e in ctx.AllBlockLevel)
        {
            if (e is Paragraph p)
            {
                var tool = ctx.GetTool(p);
                var data = tool.GetFeature(GostParagraphData.Key)!;

                if (data.StructuralElementHeader != null)
                {
                    currentElement = data.StructuralElementHeader;
                }

                if (tool.HeadingData?.Level == 1)
                {
                    currentHeading1 = tool;
                }

                data.OfStructuralElement = currentElement;
                tool.AssociatedHeading1 = currentHeading1;

                if (currentElement == null && currentHeading1 == null)
                    tool.IsIgnored = true;
            } else if (e is Table t)
            {
                if (currentElement == null && currentHeading1 == null)
                    ctx.GetTool(t).IsOutsideOfDocument = true;
            }
        }
    }
}