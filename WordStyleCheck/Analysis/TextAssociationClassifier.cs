namespace WordStyleCheck.Analysis;

public class TextAssociationClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        StructuralElement? currentElement = null;
        ParagraphPropertiesTool? currentHeading1 = null;

        foreach (var p in ctx.AllParagraphs)
        {
            var tool = ctx.GetTool(p);

            if (tool.StructuralElementHeader != null)
            {
                currentElement = tool.StructuralElementHeader;
            }
            
            if (tool.HeadingData?.Level == 1)
            {
                currentHeading1 = tool;
            }

            tool.OfStructuralElement = currentElement;
            tool.AssociatedHeading1 = currentHeading1;
        }
    }
}