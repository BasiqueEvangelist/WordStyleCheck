using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class PageSizeLint(bool allowLandscape = true) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectPageSize"];

    public void Run(LintContext ctx)
    {
        foreach (var section in ctx.Document.AllSections)
        {
            if (section.Type == SectionMarkValues.Continuous || section.Type == SectionMarkValues.NextColumn) continue;

            if (section.PageSize == null) continue;
            
            Size size = section.PageSize.Value;

            if (section.Orientation == PageOrientationValues.Landscape && allowLandscape)
                size = new Size(size.Height, size.Width);

            Size target = new(11906, 16838);

            if (Math.Abs(size.Width - target.Width) > 5 || Math.Abs(size.Height - target.Height) > 5)
            {
                ctx.AddMessage(new LintMessage("IncorrectPageSize", new SectionDiagnosticContext(section))
                {
                    Parameters = new()
                    {
                        ["ExpectedWidthCm"] = Utils.TwipsToCm(target.Width).ToString(CultureInfo.InvariantCulture),
                        ["ExpectedHeightCm"] = Utils.TwipsToCm(target.Height).ToString(CultureInfo.InvariantCulture),
                        
                        ["ActualWidthCm"] = Utils.TwipsToCm(size.Width).ToString(CultureInfo.InvariantCulture),
                        ["ActualHeightCm"] = Utils.TwipsToCm(size.Height).ToString(CultureInfo.InvariantCulture),
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