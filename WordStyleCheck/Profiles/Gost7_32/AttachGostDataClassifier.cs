using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Gost7_32;

public class AttachGostDataClassifier : IClassifier
{
    public void Classify(DocumentAnalysisContext ctx)
    {
        foreach (var p in ctx.AllParagraphs)
        {
            var tool = ctx.GetTool(p);
            tool.SetFeature(GostParagraphData.Key, new GostParagraphData(tool));
        }
    }
}