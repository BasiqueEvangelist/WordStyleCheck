using WordStyleCheck.Analysis;

namespace WordStyleCheck.Lints;

public class LintContext(DocumentAnalysisContext document, bool generateRevisions)
{
    public DocumentAnalysisContext Document { get; } = document;
    public bool GenerateRevisions { get; } = generateRevisions;

    public List<LintMessage> Messages { get; } = [];

    public Predicate<string> LintIdFilter { get; set; } = _ => true;
    
    public void AddMessage(LintMessage message)
    {
        if (!LintIdFilter(message.Id)) return;
        
        Messages.Add(message);
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