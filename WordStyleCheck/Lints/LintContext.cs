using DocumentFormat.OpenXml.Packaging;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Lints;

public class LintContext(DocumentAnalysisContext document, bool generateRevisions)
{
    public DocumentAnalysisContext Document { get; } = document;
    public bool GenerateRevisions { get; } = generateRevisions;

    public List<LintMessage> Messages { get; } = [];
    
    public void AddMessage(LintMessage message)
    {
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