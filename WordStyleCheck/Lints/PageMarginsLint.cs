using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class PageMarginsLint(PageMargins target) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectPageMargins"];

    public void Run(LintContext ctx)
    {
        foreach (var section in ctx.Document.AllSections)
        {
            if (section.Type == SectionMarkValues.Continuous || section.Type == SectionMarkValues.NextColumn) continue;

            if (section.PageMargins == null) continue;
            
            PageMargins margins = section.PageMargins.Value;

            if (section.Orientation == PageOrientationValues.Landscape)
                margins = margins.Rotate();
            
            if (!margins.CloseTo(target))
            {
                ctx.AddMessage(new LintMessage("IncorrectPageMargins", new SectionDiagnosticContext(section))
                {
                    Parameters = new()
                    {
                        ["ExpectedTopCm"] = Utils.TwipsToCm(target.Top).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedBottomCm"] = Utils.TwipsToCm(target.Bottom).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedLeftCm"] = Utils.TwipsToCm(target.Left).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedRightCm"] = Utils.TwipsToCm(target.Right).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedHeaderCm"] = Utils.TwipsToCm(target.Header).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedFooterCm"] = Utils.TwipsToCm(target.Footer).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedGutterCm"] = Utils.TwipsToCm(target.Gutter).ToString(CultureInfo.InvariantCulture),
                        
                        ["ActualTopCm"] = Utils.TwipsToCm(margins.Top).ToString(CultureInfo.InvariantCulture),
                        ["ActualBottomCm"] = Utils.TwipsToCm(margins.Bottom).ToString(CultureInfo.InvariantCulture),
                        ["ActualLeftCm"] = Utils.TwipsToCm(margins.Left).ToString(CultureInfo.InvariantCulture),
                        ["ActualRightCm"] = Utils.TwipsToCm(margins.Right).ToString(CultureInfo.InvariantCulture),
                        ["ActualHeaderCm"] = Utils.TwipsToCm(margins.Header).ToString(CultureInfo.InvariantCulture),
                        ["ActualFooterCm"] = Utils.TwipsToCm(margins.Footer).ToString(CultureInfo.InvariantCulture),
                        ["ActualGutterCm"] = Utils.TwipsToCm(margins.Gutter).ToString(CultureInfo.InvariantCulture),
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