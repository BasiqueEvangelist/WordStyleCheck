using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using Text = DocumentFormat.OpenXml.Drawing.Text;

namespace WordStyleCheck.Lints;

public class ForcePageBreakBeforeLint(Predicate<ParagraphPropertiesTool> predicate, string messageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];
    
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue;
            
            if (tool.PageBreakBefore) continue;

            bool found = false;
            foreach (var run in Utils.DirectRunChildren(p))
            {
                foreach (var child in run)
                {
                    if (child is Text t && !string.IsNullOrWhiteSpace(t.Text))
                        goto outer;
                    
                    if (child is LastRenderedPageBreak || (child is Break b && b.Type?.Value == BreakValues.Page))
                    {
                        found = true;
                        goto outer;
                    }
                }
            }
            outer: ;
            
            if (found) continue;

            OpenXmlElement? prev = p.PreviousSibling();

            while (prev != null)
            {
                if (prev is Table)
                {
                    break;
                }

                if (prev is Paragraph prevP)
                {
                    foreach (var run in Utils.DirectRunChildren(prevP))
                    {
                        foreach (var child in run.Reverse())
                        {
                            if (child is Text t && !string.IsNullOrWhiteSpace(t.Text))
                                goto outer2;

                            if (child is LastRenderedPageBreak ||
                                (child is Break b && b.Type?.Value == BreakValues.Page))
                            {
                                found = true;
                                goto outer2;
                            }
                        }
                    }

                    if (ctx.Document.GetTool(prevP).PageBreakBefore)
                    {
                        found = true;
                        break;
                    }
                }

                prev = prev.PreviousSibling();
            }
            outer2:
            if (found || prev == null) continue;
            
            ctx.AddMessage(new LintMessage(messageId, new ParagraphDiagnosticContext(p)));
        }
    }
}