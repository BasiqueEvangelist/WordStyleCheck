using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class PageMarginsLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var section in ctx.Document.AllSections)
        {
            if (section.Type == SectionMarkValues.Continuous || section.Type == SectionMarkValues.NextColumn) continue;

            if (section.PageMargins == null) continue;
            
            PageMargins margins = section.PageMargins.Value;

            if (section.Orientation == PageOrientationValues.Landscape)
                margins = margins.Rotate();

            var target = new PageMargins(1134, 1134, 1701, 851, 709, 709, 0);
            
            if (!margins.CloseTo(target))
            {
                ctx.AddMessage(new LintMessage("IncorrectPageMargins", new SectionDiagnosticContext(section))
                {
                    Parameters = new()
                    {
                        ["ExpectedTopCm"] = Utils.TwipsToCm(target.Top).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedBottomCm"] = Utils.TwipsToCm(target.Bottom).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedLeftCm"] = Utils.TwipsToCm(target.Left).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedRightCm"] = Utils.TwipsToCm(target.Right).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedHeaderCm"] = Utils.TwipsToCm(target.Header).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedFooterCm"] = Utils.TwipsToCm(target.Footer).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedGutterCm"] = Utils.TwipsToCm(target.Gutter).ToString(CultureInfo.CurrentCulture),
                        
                        ["ActualTopCm"] = Utils.TwipsToCm(margins.Top).ToString(CultureInfo.CurrentCulture),
                        ["ActualBottomCm"] = Utils.TwipsToCm(margins.Bottom).ToString(CultureInfo.CurrentCulture),
                        ["ActualLeftCm"] = Utils.TwipsToCm(margins.Left).ToString(CultureInfo.CurrentCulture),
                        ["ActualRightCm"] = Utils.TwipsToCm(margins.Right).ToString(CultureInfo.CurrentCulture),
                        ["ActualHeaderCm"] = Utils.TwipsToCm(margins.Header).ToString(CultureInfo.CurrentCulture),
                        ["ActualFooterCm"] = Utils.TwipsToCm(margins.Footer).ToString(CultureInfo.CurrentCulture),
                        ["ActualGutterCm"] = Utils.TwipsToCm(margins.Gutter).ToString(CultureInfo.CurrentCulture),
                    },
                    AutoFix = () =>
                    {
                        // TODO: add generate-revisions support.
                        var pgMar = section.Properties.GetOrAddFirstChild<PageMargin>();
                        
                        pgMar.Top = target.Top;
                        pgMar.Bottom = target.Bottom;
                        pgMar.Left = (uint) target.Left;
                        pgMar.Right = (uint) target.Right;
                        pgMar.Header = (uint) target.Header;
                        pgMar.Footer = (uint) target.Footer;
                        pgMar.Gutter = (uint) target.Gutter;
                    }
                });
            }
        }
    }
}