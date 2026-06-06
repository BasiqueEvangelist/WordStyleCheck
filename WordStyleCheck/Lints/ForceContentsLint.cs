using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ForceContentsLint(Predicate<ParagraphPropertiesTool> predicate, Func<ParagraphPropertiesTool, string> forcedContents, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (!predicate(tool)) continue;
            
            string proper = forcedContents(tool);
            
            if (tool.Contents == proper) continue;

            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = proper,
                        ["Actual"] = tool.Contents
                    }
                });
            }
            else
            {
                ctx.MarkAutoFixed();

                var props = (RunProperties?) Utils.DirectRunChildren(p).MaxBy(x => ctx.Document.GetTool(x).Contents)?.RunProperties?.CloneNode(true);
                
                // TODO: add proper support for generate-revisions.
                foreach (var child in p.ChildElements.ToList())
                {
                    if (child is ParagraphProperties) continue;

                    child.Remove();
                }

                p.Append(new Run(new Text(proper))
                {
                    RunProperties = props
                });
            }
        }
    }
}