using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class TocReferencesLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["NoToc", "ShouldNotBeInToc", "ShouldBeInToc"];
    
    public void Run(LintContext ctx)
    {
        var tocParagraphs = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Where(x => x.IsTableOfContents)
            .ToList();

        if (tocParagraphs.Count == 0)
        {
            ctx.AddMessage(new LintMessage("NoToc", new StartOfDocumentDiagnosticContext(ctx.Document.AllParagraphs.First())));
            
            return;
        }

        HashSet<Paragraph> referencedParagraphs = [];

        foreach (var p in tocParagraphs)
        {
            var links = p.Paragraph.Descendants<Hyperlink>().ToList();

            if (links.Count != 1)
            {
                // Probably some random junk.
                // TODO: check for this.
                continue;
            }
            
            if (links[0].Anchor?.Value is not {} anchor) continue;
            
            if (!ctx.Document.BookmarkStarts.TryGetValue(anchor, out var bookmarkStart)) continue;

            var target = Utils.AscendToAnscestor<Paragraph>(bookmarkStart);
            
            if (target == null) continue; // ???

            var targetTool = ctx.Document.GetTool(target);

            if (targetTool is not ({Class: ParagraphClass.Heading, HeadingData.Level: < 4} or {StructuralElementHeader: StructuralElement.Introduction or StructuralElement.Conclusion or StructuralElement.Bibliography or StructuralElement.Appendix}) && targetTool.OfStructuralElement != StructuralElement.Appendix)
            {
                ctx.AddMessage(new LintMessage("ShouldNotBeInToc", new ParagraphDiagnosticContext(p.Paragraph)));
                continue;
            }

            referencedParagraphs.Add(target);
        }

        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool is not ({Class: ParagraphClass.Heading, HeadingData.Level: < 4} or {StructuralElementHeader: StructuralElement.Introduction or StructuralElement.Conclusion or StructuralElement.Bibliography or StructuralElement.Appendix})) continue;
            
            if (referencedParagraphs.Contains(tool.Paragraph)) continue;
            
            ctx.AddMessage(new LintMessage("ShouldBeInToc", new ParagraphDiagnosticContext(p)));

        }
    }
}