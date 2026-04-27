using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectHeadingTextLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectHeadingText"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.HeadingData == null) continue;
            if (tool.HeadingData.IsConclusion) continue;

            string text = tool.Contents;
            
            string correct;
            if (tool.OfNumbering is NumberingPropertiesTool)
                correct = tool.HeadingData.Title;
            else
                correct = tool.HeadingData.Number + " " + tool.HeadingData.Title;

            if (text != correct)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("IncorrectHeadingText", DiagnosticType.FormattingError,
                        new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = correct,
                            ["Actual"] = text
                        }
                    });
                }
                else
                {
                    ctx.MarkAutoFixed();

                    foreach (var child in p.ChildElements.ToList())
                    {
                        if (child is ParagraphProperties) continue;
                            
                        child.Remove();
                    }
                        
                    p.Append(new Run(new Text(correct)));
                }
            }
        }
    }
}