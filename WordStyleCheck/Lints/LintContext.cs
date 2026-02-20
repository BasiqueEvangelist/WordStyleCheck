using DocumentFormat.OpenXml.Packaging;

namespace WordStyleCheck.Lints;

public class LintContext(WordprocessingDocument document, bool autofixEnabled, bool generateRevisions)
{
    public WordprocessingDocument Document { get; } = document;
    public bool AutofixEnabled { get; } = autofixEnabled;
    public bool GenerateRevisions { get; } = generateRevisions;

    public List<LintMessage> Messages { get; } = [];
    
    private bool _documentChanged = false;
    public bool DocumentChanged => _documentChanged;
    public void MarkDocumentChanged()
    {
        if (!AutofixEnabled)
        {
            throw new InvalidOperationException("Cannot change document if automatic fixing is disabled");
        }
        
        _documentChanged = true;
    }

    public void AddMessage(LintMessage message)
    {
        Messages.Add(message);
    }
}