using WordStyleCheck.Analysis;

namespace WordStyleCheck;

public interface ILintContext
{
    public DocumentAnalysisContext Document { get; }
    public bool GenerateRevisions { get; }

    public void AddMessage(LintDiagnostic diagnostic);
}