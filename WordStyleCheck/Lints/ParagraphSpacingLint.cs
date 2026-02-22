using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphSpacingLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            // TODO: enforce this for table cell content, headers, captions.
            if (tool.Class != ParagraphPropertiesTool.ParagraphClass.BodyText) continue; 
            
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