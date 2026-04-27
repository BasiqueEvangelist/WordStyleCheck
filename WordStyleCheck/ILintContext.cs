using WordStyleCheck.Analysis;

namespace WordStyleCheck;

public interface ILintContext
{
    public DocumentAnalysisContext Document { get; }
    public bool AutomaticallyFix { get; }
    public bool GenerateRevisions { get; }

    public void AddMessage(LintDiagnostic diagnostic);
    public void MarkAutoFixed();
}