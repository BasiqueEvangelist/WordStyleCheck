using DocumentFormat.OpenXml.Packaging;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Lints;

public class LintContext(DocumentAnalysisContext document, bool generateRevisions)
{
    public DocumentAnalysisContext Document { get; } = document;
    public bool GenerateRevisions { get; } = generateRevisions;

    public List<LintMessage> Messages { get; } = [];

    public Predicate<LintMessage> LintFilter { get; set; } = _ => true;
    
    public void AddMessage(LintMessage message)
    {
        if (!LintFilter(message)) return;
        
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