using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectHeadingTextLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.HeadingData == null) continue;

            string text = Utils.CollectParagraphText(p);
            
            string correct;
            if (tool.OfNumbering is NumberingPropertiesTool)
                correct = tool.HeadingData.Title;
            else
                correct = tool.HeadingData.Number + " " + tool.HeadingData.Title;

            if (text != correct)
            {
                ctx.AddMessage(new LintMessage("IncorrectHeadingText", new ParagraphDiagnosticContext(p))
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