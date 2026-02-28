using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphLineSpacingLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            if (tool.IsEmptyOrDrawing) continue;
            
            // TODO: enforce this for table cell content, headers, captions.
            if (tool.Class != ParagraphClass.BodyText) continue;
            
            // TODO: enforce this for numberings.
            if (tool.OfNumbering != null) continue;
            
            // Appendices can have weird formatting.
            if (tool.OfStructuralElement == StructuralElement.Appendix) continue;
            
            if (tool.LineSpacing != 360)
            {
                ctx.AddMessage(new LintMessage(
                    "IncorrectParagraphLineSpacing",
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                    
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.SpacingBetweenLines == null)
                                p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();

                            p.ParagraphProperties.SpacingBetweenLines.Line = "360";
                        }
                    }
                );
            }
        }
    }
}