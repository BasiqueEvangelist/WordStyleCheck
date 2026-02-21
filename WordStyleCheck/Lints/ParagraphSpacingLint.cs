using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphSpacingLint : ILint
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

            ParagraphPropertiesTool tool = ParagraphPropertiesTool.Get(ctx.Document, p);
            
            if (tool.ContainingTableCell != null) continue; // TODO: enforce this for table cell content.
            if (tool.OutlineLevel != null) continue; // TODO: enforce this for headers
            if (tool.ProbablyCaption) continue; // TODO: enforce this for captions
            
            if (tool.BeforeSpacing != 120 || tool.LineSpacing != 360)
            {
                ctx.AddMessage(new LintMessage(
                    "Paragraph doesn't have set before and line spacing",
                    new ParagraphDiagnosticContext(p))
                    {
                        Values = new("120 and 360", $"{tool.BeforeSpacing} and {tool.LineSpacing}"),
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                    
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.SpacingBetweenLines == null)
                                p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();

                            p.ParagraphProperties.SpacingBetweenLines.Before = "120";
                            p.ParagraphProperties.SpacingBetweenLines.Line = "360";
                        }
                    }
                );
            }
        }
    }
}