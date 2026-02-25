using System.Globalization;
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
            if (tool.Class != ParagraphClass.BodyText) continue;
            
            // Appendices can have weird formatting.
            if (tool.OfStructuralElement == StructuralElement.Appendix) continue;
            
            if (tool.BeforeSpacing != 120 || tool.LineSpacing != 360 || tool.AfterSpacing != 120)
            {
                ctx.AddMessage(new LintMessage(
                    "IncorrectParagraphSpacing",
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["ExpectedBeforeCm"] = Utils.TwipsToCm(120).ToString(CultureInfo.CurrentCulture),
                            ["ExpectedLineCm"] = Utils.TwipsToCm(360).ToString(CultureInfo.CurrentCulture),
                            ["ExpectedAfterCm"] = Utils.TwipsToCm(120).ToString(CultureInfo.CurrentCulture),
                            
                            ["ActualBeforeCm"] = Utils.TwipsToCm(tool.BeforeSpacing ?? 0).ToString(CultureInfo.CurrentCulture),
                            ["ActualLineCm"] = Utils.TwipsToCm(tool.LineSpacing ?? 0).ToString(CultureInfo.CurrentCulture),
                            ["ActualAfterCm"] = Utils.TwipsToCm(tool.AfterSpacing ?? 0).ToString(CultureInfo.CurrentCulture),
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                    
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.SpacingBetweenLines == null)
                                p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();

                            p.ParagraphProperties.SpacingBetweenLines.Before = "120";
                            p.ParagraphProperties.SpacingBetweenLines.Line = "360";
                            p.ParagraphProperties.SpacingBetweenLines.After = "120";
                        }
                    }
                );
            }
        }
    }
}