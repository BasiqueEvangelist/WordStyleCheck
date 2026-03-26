using Avalonia.Media;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

using AvaloniaRun = Avalonia.Controls.Documents.Run;

namespace WscDebugTool.ViewModels;

public class RunViewModel(DocumentFileViewModel doc, ParagraphViewModel paragraph, Run run) : ViewModelBase
{
    public RunPropertiesTool Tool { get; } = doc.Document.GetTool(run);

    public string Contents => Tool.Contents;

    public AvaloniaRun PrepareRun()
    {
        AvaloniaRun run = new AvaloniaRun(Tool.Contents.Replace("\t", "    "));

        if (Tool.Bold) run.FontWeight = FontWeight.Bold;
        if (Tool.Italic) run.FontStyle = FontStyle.Italic;

        return run;
    }
}