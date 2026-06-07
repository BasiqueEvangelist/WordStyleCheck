using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class ForbidPageBreaksLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["PageBreakForbidden"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            bool bad = false;
            
            foreach (var run in Utils.DirectRunChildren(p))
            {
                foreach (var child in run)
                {
                    if (child is Break b && b.Type?.Value == BreakValues.Page)
                    {
                        bad = true;
                        goto outer;
                    }
                }
            }
            outer: ;
            
            if (tool.PageBreakBefore || bad)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic("PageBreakForbidden", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
                }
                else
                {
                    ctx.MarkAutoFixed();

                    if (tool.PageBreakBefore)
                    {
                        p.ParagraphProperties ??= new ParagraphProperties();
                        if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                        p.ParagraphProperties.PageBreakBefore ??= new PageBreakBefore();
                        p.ParagraphProperties.PageBreakBefore.Val = false;
                    }

                    if (bad)
                    {
                        // TODO: generate-revisions.
                        
                        foreach (var run in Utils.DirectRunChildren(p))
                        {
                            foreach (var child in run.ChildElements.ToList())
                            {
                                if (child is Break b && b.Type?.Value == BreakValues.Page)
                                {
                                    child.Remove();
                                }
                            }
                        }
                    }
                }
            }
        }

        foreach (var section in ctx.Document.AllSections)
        {
            int idx = ((List<SectionPropertiesTool>)ctx.Document.AllSections).IndexOf(section);
            
            if (idx == ctx.Document.AllSections.Count - 1 ) continue;

            if (section.Type != SectionMarkValues.NextPage && section.Type != SectionMarkValues.EvenPage &&
                section.Type != SectionMarkValues.OddPage) continue;
            
            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic("PageBreakForbidden", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(section.Paragraphs[^1])));
            }
            else
            {
                ctx.MarkAutoFixed();

                // TODO: generate-revisions.
                
                section.Paragraphs[^1].ParagraphProperties!.SectionProperties!.Remove();
            }
        }
    }
}