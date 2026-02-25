using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectCaptionTextLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData == null) continue;

            string text = Utils.CollectParagraphText(p);
            string correct = tool.CaptionData.Value.GetCorrectText(text);

            if (text != correct)
            {
                ctx.AddMessage(new LintMessage("IncorrectCaptionText", new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = correct,
                        ["Actual"] = text
                    },
                    // TODO: figure out how to replace paragraph text properly.
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