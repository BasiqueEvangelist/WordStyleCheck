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
                        // TODO!!!: improve this.
                        ["Expected"] = target.ToString(),
                        ["Actual"] = margins.ToString()
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