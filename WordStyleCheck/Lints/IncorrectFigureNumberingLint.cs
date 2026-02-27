using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectFigureNumberingLint : ILint
{
    public void Run(LintContext ctx)
    {
        var figures = ctx.Document.AllParagraphs
            .Select(x => ctx.Document.GetTool(x))
            .Where(x => x is
            {
                CaptionData: {Type: CaptionType.Figure},
                OfStructuralElement: not StructuralElement.Appendix // TODO: handle this for figures in appendices too
            })
            .ToList();

        for (int i = 0; i < figures.Count; i++)
        {
            string correctNumber = (i + 1).ToString();

            if (figures[i].CaptionData!.Value.Number != correctNumber)
            {
                ctx.AddMessage(new LintMessage("IncorrectFigureNumbering", new ParagraphDiagnosticContext(figures[i].Paragraph))
                {
                    Parameters = new()
                    {
                        ["Expected"] = correctNumber,
                        ["Actual"] = figures[i].CaptionData!.Value.Number
                    }
                });
            }
        }
    }
}