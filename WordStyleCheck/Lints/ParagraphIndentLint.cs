using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphIndentLint : ILint
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
            if (tool.Class != ParagraphClass.BodyText) continue; 
            
            // Appendices can have weird formatting.
            if (tool.OfStructuralElement == StructuralElement.Appendix) continue;
            
            if (tool.FirstLineIndent != 709 || tool.LeftIndent is not null and not 0)
            {
                ctx.AddMessage(new LintMessage(
                    "IncorrectParagraphIndent",
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["ExpectedFirstLine"] = Utils.TwipsToCm(709).ToString(CultureInfo.CurrentCulture),
                            ["ExpectedLeft"] = "0",
                            ["ActualFirstLine"] = Utils.TwipsToCm(tool.FirstLineIndent ?? 0).ToString(CultureInfo.CurrentCulture),
                            ["ActualLeft"] = Utils.TwipsToCm(tool.LeftIndent ?? 0).ToString(CultureInfo.CurrentCulture),
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.Indentation == null)
                                p.ParagraphProperties.Indentation = new Indentation();

                            p.ParagraphProperties.Indentation.Left = "0";
                            p.ParagraphProperties.Indentation.FirstLine = "709";
                        }
                    }
                );
            }
        }
    }
}