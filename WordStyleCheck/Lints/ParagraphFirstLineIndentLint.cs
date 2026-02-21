using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Lints;

public class ParagraphFirstLineIndentLint : ILint
{
    public void Run(LintContext ctx)
    {
        var body = ctx.Document.MainDocumentPart?.Document?.Body;

        if (body == null) return;
        
        foreach (var p in body.Descendants<Paragraph>())
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool tool = new(ctx.Document, p);
            
            if (tool.FirstLineIndent != 709)
            {
                ctx.AddMessage(new LintMessage($"Paragraph didn't have set first line indent (expected 709, was {tool.FirstLineIndent})", ctx.AutofixEnabled, Context.FromParagraph(p)));

                if (ctx.AutofixEnabled)
                {
                    if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                    
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    if (p.ParagraphProperties.Indentation == null)
                        p.ParagraphProperties.Indentation = new Indentation();

                    p.ParagraphProperties.Indentation.FirstLine = "709";
                    ctx.MarkDocumentChanged();
;                }
            }
        }
    }
}