using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck;
using WordStyleCheck.Analysis;

namespace WscDebugTool.ViewModels;

public class ParagraphViewModel(DocumentFileViewModel doc, Paragraph p) : ViewModelBase
{
    public ParagraphPropertiesTool Tool { get; } = doc.Document.GetTool(p);

    public List<RunViewModel> Runs => Utils.DirectRunChildren(p).Select(x => new RunViewModel(doc, this, x)).ToList();

    public void Focus()
    {
        doc.Focus(p);
    }
}