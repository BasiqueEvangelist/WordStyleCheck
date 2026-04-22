namespace WordStyleCheck.Analysis;

public interface IClassifier
{
    void Classify(DocumentAnalysisContext ctx);
}