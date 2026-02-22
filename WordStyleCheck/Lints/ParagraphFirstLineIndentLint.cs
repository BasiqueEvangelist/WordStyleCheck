using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphFirstLineIndentLint : ILint
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
            
            // TODO: enforce this for captions, heading, table content.
            if (tool.Class != ParagraphPropertiesTool.ParagraphClass.BodyText) continue; 
            
            if (tool.FirstLineIndent != 709)
            {
                ctx.AddMessage(new LintMessage(
                    "Paragraph didn't have set first line indent",
                    new ParagraphDiagnosticContext(p))
                    {
                        Values = new("709", tool.FirstLineIndent != null ? tool.FirstLineIndent.ToString() : null),
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.Indentation == null)
                                p.ParagraphProperties.Indentation = new Indentation();

                            p.ParagraphProperties.Indentation.FirstLine = "709";
                        }
                    }
                );
            }
        }
    }
}