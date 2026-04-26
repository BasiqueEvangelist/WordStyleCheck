using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectCaptionTextLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectCaptionText"];

    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData == null) continue;

            string text = tool.Contents.Trim();
            string correct = tool.CaptionData.Value.GetCorrectText(text);

            if (text != correct)
            {
                ctx.AddMessage(new LintDiagnostic("IncorrectCaptionText", new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = correct,
                        ["Actual"] = text
                    },
                    AutoFix = () =>
                    {
                        foreach (var child in p.ChildElements.ToList())
                        {
                            if (child is ParagraphProperties) continue;
                            
                            child.Remove();
                        }
                        
                        p.Append(new Run(new Text(correct)));
                    }
                });
            }
        }
    }
}