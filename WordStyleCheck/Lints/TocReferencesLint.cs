using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class TocReferencesLint(Predicate<ParagraphPropertiesTool> shouldBeInToc) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["NoToc", "ShouldNotBeInToc", "ShouldBeInToc"];
    
    public void Run(ILintContext ctx)
    {
        var tocParagraphs = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Where(x => x.IsTableOfContents)
            .ToList();

        if (tocParagraphs.Count == 0)
        {
            ctx.AddMessage(new LintDiagnostic("NoToc", DiagnosticType.ContentError, new StartOfDocumentDiagnosticContext()));
            
            return;
        }

        HashSet<Paragraph> referencedParagraphs = [];

        foreach (var p in tocParagraphs)
        {
            var links = p.Paragraph.Descendants<Hyperlink>().Select(x => x.Anchor?.Value).Where(x => x is not null).ToList();

            string anchor;
            
            if (links.Count != 1)
            {
                var anchors = p.Paragraph
                    .Descendants<FieldCode>()
                    .Select(x => x.Text.Trim())
                    .Where(x => x.StartsWith("PAGEREF "))
                    .Select(x => x.Substring("PAGEREF".Length).Trim())
                    .ToList();
                
                if (anchors.Count != 1)
                {
                    // Probably some random junk.
                    // TODO: check for this.
                    continue;
                }

                anchor = anchors[0];
                
                if (anchor.EndsWith("\\h")) anchor = anchor[..^2];
            }
            else
            {
                anchor = links[0]!;
            }

            if (!ctx.Document.BookmarkStarts.TryGetValue(anchor, out var bookmarkStart)) continue;

            var target = Utils.AscendToAnscestor<Paragraph>(bookmarkStart);
            
            if (target == null) continue; // ???

            var targetTool = ctx.Document.GetTool(target);

            referencedParagraphs.Add(target);

            if (!shouldBeInToc(targetTool))
            {
                ctx.AddMessage(new LintDiagnostic("ShouldNotBeInToc", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p.Paragraph)));
                continue;
            }
        }

        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!shouldBeInToc(tool)) continue;
            
            if (referencedParagraphs.Contains(tool.Paragraph)) continue;
            
            ctx.AddMessage(new LintDiagnostic("ShouldBeInToc", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));

        }
    }
}