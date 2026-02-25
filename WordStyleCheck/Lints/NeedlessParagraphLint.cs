using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class NeedlessParagraphLint : ILint
{
    public void Run(LintContext ctx)
    {
        var body = ctx.Document.Document.MainDocumentPart?.Document?.Body;

        if (body == null) return;

        var paragraphs = body.ChildElements.OfType<Paragraph>().ToList();
        
        for (int i = paragraphs.Count - 1; i >= 1; i--)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(paragraphs[i], 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool prevTool = ctx.Document.GetTool(paragraphs[i - 1]);
            ParagraphPropertiesTool curTool = ctx.Document.GetTool(paragraphs[i]);

            if (prevTool.OutlineLevel != null || curTool.OutlineLevel != null) continue;
            if (prevTool.Class != ParagraphClass.BodyText
             || curTool.Class != ParagraphClass.BodyText)
                continue;

            var prevParagraphText = Utils.CollectParagraphText(paragraphs[i - 1]).TrimEnd();
            
            if (prevParagraphText.Length == 0 || prevParagraphText[^1] == '.' ||prevParagraphText[^1] == '?' || prevParagraphText[^1] == '!')
                continue;

            var paraText = Utils.CollectParagraphText(paragraphs[i]);
            
            if (char.IsUpper(paraText[0]))
                continue;
            
            var p = paragraphs[i];
            var prev = paragraphs[i - 1];
            
            ctx.AddMessage(new LintMessage("NeedlessParagraphBreak", new MergeParagraphsDiagnosticContext(paragraphs[i - 1], paragraphs[i]))
            {
                AutoFix = () =>
                {
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
                    

                    if (ctx.GenerateRevisions)
                    {
                        if (prev.ParagraphProperties == null)
                            prev.ParagraphProperties = new ParagraphProperties();

                        if (prev.ParagraphProperties.ParagraphMarkRunProperties == null)
                            prev.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties();

                        if (prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted == null)
                            prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted = new Deleted();

                        Utils.StampTrackChange(prev.ParagraphProperties.ParagraphMarkRunProperties.Deleted);
                    }
                    else
                    {
                        p.ParagraphProperties = null;

                        foreach (var child in p.ChildElements)
                        {
                            prev.Append(child.CloneNode(true));
                        }
                        
                        p.Remove();
                    }
                }
            });
        }
    }
}