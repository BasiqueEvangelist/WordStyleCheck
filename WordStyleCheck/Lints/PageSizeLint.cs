using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class PageSizeLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var section in ctx.Document.AllSections)
        {
            if (section.Type == SectionMarkValues.Continuous || section.Type == SectionMarkValues.NextColumn) continue;

            if (section.PageSize == null) continue;
            
            Size size = section.PageSize.Value;

            if (section.Orientation == PageOrientationValues.Landscape)
                size = new Size(size.Height, size.Width);

            Size target = new(11906, 16838);

            if (size != target)
            {
                ctx.AddMessage(new LintMessage("IncorrectPageSize", new SectionDiagnosticContext(section))
                {
                    Parameters = new()
                    {
                        ["ExpectedWidthCm"] = Utils.TwipsToCm(target.Width).ToString(CultureInfo.CurrentCulture),
                        ["ExpectedHeightCm"] = Utils.TwipsToCm(target.Height).ToString(CultureInfo.CurrentCulture),
                        
                        ["ActualWidthCm"] = Utils.TwipsToCm(size.Width).ToString(CultureInfo.CurrentCulture),
                        ["ActualHeightCm"] = Utils.TwipsToCm(size.Height).ToString(CultureInfo.CurrentCulture),
                    },
                    AutoFix = () =>
                    {
                        // TODO: add generate-revisions support.
                        var pgSz = section.Properties.GetOrAddFirstChild<PageSize>();
                        
                        pgSz.Width = 11906;
                        pgSz.Height = 16838;
                    }
                });
            }
        }
    }
}