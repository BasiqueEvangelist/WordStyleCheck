using System.Drawing;
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

            if (size != new Size(11906, 16838))
            {
                ctx.AddMessage(new LintMessage("Page size must be A4", new SectionDiagnosticContext(section))
                {
                    Values = new("11906x16838", $"{size.Width}x{size.Height}"),
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