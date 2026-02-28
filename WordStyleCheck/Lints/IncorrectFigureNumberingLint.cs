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

        int underHeadingNumber = 0;
        bool prevCorrectWasWithSection = false;
        for (int i = 0; i < figures.Count; i++)
        {
            string correctNumber = (i + 1).ToString();

            underHeadingNumber += 1;

            if (i > 0 && figures[i - 1].AssociatedHeading1 != figures[i].AssociatedHeading1) underHeadingNumber = 1;

            var actualNumber = figures[i].CaptionData!.Value.Number;

            if (actualNumber == correctNumber)
            {
                prevCorrectWasWithSection = false;
                continue;
            }

            string? correctNumberSection = null;
            if (figures[i].AssociatedHeading1?.HeadingData?.Number is {} headingNumber)
            {
                correctNumberSection = $"{headingNumber}.{underHeadingNumber}";

                if (actualNumber == correctNumberSection)
                {
                    prevCorrectWasWithSection = true;
                    continue;
                }
            }

            ctx.AddMessage(new LintMessage("IncorrectFigureNumbering", new ParagraphDiagnosticContext(figures[i].Paragraph))
            {
                Parameters = new()
                {
                    ["Expected"] = prevCorrectWasWithSection && correctNumberSection != null ? correctNumberSection : correctNumber,
                    ["Actual"] = figures[i].CaptionData!.Value.Number
                }
            });
        }
    }
}