using WordStyleCheck.Analysis;

namespace WordStyleCheck.Lints;

public class LintContext(DocumentAnalysisContext document, bool generateRevisions)
{
    public DocumentAnalysisContext Document { get; } = document;
    public bool GenerateRevisions { get; } = generateRevisions;

    public List<LintDiagnostic> Messages { get; } = [];

    public Predicate<string> LintIdFilter { get; set; } = _ => true;
    
    public void AddMessage(LintDiagnostic diagnostic)
    {
        if (!LintIdFilter(diagnostic.Id)) return;
        
        Messages.Add(diagnostic);
    }

    public bool RunAllAutoFixes()
    {
        bool changed = false;
        
        foreach (var message in Messages)
        {
            if (message.AutoFix != null)
            {
                message.AutoFix();
                changed = true;
            }
        }

        return changed;
    }
}