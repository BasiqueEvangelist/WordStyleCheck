using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class InterParagraphSpacingLint(List<InterParagraphSpacingLint.SpacingEntry> entries, string messageId) : ILint
{
    public void Run(LintContext ctx)
    {
        var paragraphs = ctx.Document.Document.MainDocumentPart!.Document!.Body!.ChildElements.OfType<Paragraph>().ToList();

        for (int i = 1; i < paragraphs.Count; i++)
        {
            if (paragraphs[i - 1].NextSibling() != paragraphs[i]) continue;
            
            var tool1 = ctx.Document.GetTool(paragraphs[i - 1]);
            var tool2 = ctx.Document.GetTool(paragraphs[i]);
            
            // TODO: Handle empty paragraphs later.
            if (tool1.IsEmptyOrDrawing)
            {
                continue;
            }
            
            if (tool2.IsEmptyOrDrawing)
            {
                continue;
            }
            
            if (paragraphs[i - 1].Descendants<Break>().Any()) continue;

            var entry1 = entries.FirstOrDefault(x => x.p(tool1));
            var entry2 = entries.FirstOrDefault(x => x.p(tool2));

            if (entry1 == null || entry2 == null) continue;

            int totalTwips = entry1.twipsAfter + entry2.twipsBefore; 
            
            if (entry1 == entry2 && entry1.contextualSpacing)
            {
                totalTwips = 0;
            }
            
            if ((tool1.ActualAfterSpacing ?? 0) + (tool2.ActualBeforeSpacing ?? 0) != totalTwips)
            {
                ctx.AddMessage(new LintMessage(messageId, new MergeParagraphsDiagnosticContext(paragraphs[i - 1], paragraphs[i]))
                {
                    Parameters = new()
                    {
                        ["ExpectedPt"] = (totalTwips / 20.0).ToString(),
                        ["ActualPt"] = (((tool1.ActualAfterSpacing ?? 0) + (tool2.ActualBeforeSpacing ?? 0)) / 20.0).ToString()
                    }
                });
            }
        }
    }

    public record SpacingEntry(
        Predicate<ParagraphPropertiesTool> p,
        int twipsBefore,
        int twipsAfter,
        bool contextualSpacing = false);
}