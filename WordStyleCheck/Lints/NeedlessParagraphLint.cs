using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class NeedlessParagraphLint : ILint
{
    public void Run(LintContext ctx)
    {
        var body = ctx.Document.MainDocumentPart?.Document?.Body;

        if (body == null) return;

        var paragraphs = body.ChildElements.OfType<Paragraph>().ToList();
        
        for (int i = 1; i < paragraphs.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(paragraphs[i], 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool prevTool = ParagraphPropertiesTool.Get(ctx.Document, paragraphs[i - 1]);
            ParagraphPropertiesTool curTool = ParagraphPropertiesTool.Get(ctx.Document, paragraphs[i]);

            if (prevTool.OutlineLevel != null || curTool.OutlineLevel != null) continue;
            if (prevTool.IsTableOfContents || curTool.IsTableOfContents) continue;
            if (prevTool.Style?.StyleId != curTool.Style?.StyleId) continue;

            var prevParagraphText = Utils.CollectParagraphText(paragraphs[i - 1]);
            
            if (prevParagraphText.Length == 0 || prevParagraphText[^1] == '.' ||prevParagraphText[^1] == '?' || prevParagraphText[^1] == '!')
                continue;

            var paraText = Utils.CollectParagraphText(paragraphs[i]);
            
            if (char.IsUpper(paraText[0]))
                continue;
            
            ctx.AddMessage(new LintMessage("Needless paragraph break", ctx.AutofixEnabled, new MergeParagraphsDiagnosticContext(paragraphs[i - 1], paragraphs[i])));

            if (ctx.AutofixEnabled)
            {
                var prev = paragraphs[i - 1];
                
                if (!ctx.GenerateRevisions)
                {
                    Console.WriteLine("TODO: add non-generate-revisions support for NeedlessParagraphLint");
                }
                
                if (prev.ParagraphProperties == null)
                    prev.ParagraphProperties = new ParagraphProperties();
                
                if (prev.ParagraphProperties.ParagraphMarkRunProperties == null)
                    prev.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties();
                
                if (prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted == null)
                    prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted = new Deleted();
                
                Utils.StampTrackChange(prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted);

                Paragraph p = paragraphs[i];
                
                if (!char.IsWhiteSpace(prevParagraphText[^1]) && !char.IsWhiteSpace(paraText[0]))
                {
                    Run r = new Run();
                    r.Append(new Text(" ")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });

                    if (ctx.GenerateRevisions)
                    {
                        InsertedRun ir = new InsertedRun();
                        Utils.StampTrackChange(ir);
                        ir.Append(r);

                        if (p.ParagraphProperties != null)
                            p.InsertAfter(ir, p.ParagraphProperties);
                        else
                            p.PrependChild(ir);
                    }
                    else
                    {
                        if (p.ParagraphProperties != null)
                            p.InsertAfter(r, p.ParagraphProperties);
                        else
                            p.PrependChild(r);
                    }
                }
                
                ctx.MarkDocumentChanged();
            }
        }
    }
}