using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectCaptionedNumberingLint(CaptionType type, string messageId, string mixMessageId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];

    public void Run(LintContext ctx)
    {
        var elements = ctx.Document.AllParagraphs
            .Select(x => ctx.Document.GetTool(x))
            .Where(x => x is
            {
                CaptionData: { IsContinuation: false },
                OfStructuralElement: not StructuralElement.Appendix // TODO: handle this for figures in appendices too
            } && x.CaptionData.Value.Type == type)
            .ToList();

        int underHeadingNumber = 0;
        bool? hierarchicalNumbering = null;
        for (int i = 0; i < elements.Count; i++)
        {
            string correctNumber = (i + 1).ToString();

            underHeadingNumber += 1;

            if (i > 0 && elements[i - 1].AssociatedHeading1 != elements[i].AssociatedHeading1) underHeadingNumber = 1;

            var actualNumber = elements[i].CaptionData!.Value.Number;

            string? correctNumberSection = null;
            if (elements[i].AssociatedHeading1?.HeadingData?.Number is {} headingNumber)
            {
                correctNumberSection = $"{headingNumber}.{underHeadingNumber}";
            }
            
            if (actualNumber == correctNumber)
            {
                if (hierarchicalNumbering == true)
                {
                    ctx.AddMessage(new LintMessage(mixMessageId, new ParagraphDiagnosticContext(elements[i].Paragraph))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = correctNumberSection!,
                            ["Actual"] = actualNumber
                        }
                    });
                    continue;                    
                }
                
                hierarchicalNumbering = false;
                continue;
            }
            
            if (actualNumber == correctNumberSection)
            {
                if (hierarchicalNumbering == false)
                {
                    ctx.AddMessage(new LintMessage(mixMessageId, new ParagraphDiagnosticContext(elements[i].Paragraph))
                    {
                        Parameters = new()
                        {
                            ["Expected"] = correctNumber,
                            ["Actual"] = actualNumber
                        }
                    });
                    continue;
                }

                hierarchicalNumbering = true;
                
                continue;
            }

            var actual = elements[i].CaptionData!.Value.Number;

            ctx.AddMessage(new LintMessage(messageId, new ParagraphDiagnosticContext(elements[i].Paragraph))
            {
                Parameters = new()
                {
                    ["Expected"] = (hierarchicalNumbering ?? actual.Contains(".")) && correctNumberSection != null ? correctNumberSection : correctNumber,
                    ["Actual"] = actual
                }
            });
        }
    }
}