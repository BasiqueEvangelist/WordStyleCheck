using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectCaptionedNumberingLint(Predicate<ParagraphPropertiesTool> predicate, CaptionType type, string messageId, string? mixMessageId, bool allowHierarchical = true) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [messageId];

    public void Run(ILintContext ctx)
    {
        var elements = ctx.Document.AllParagraphs
            .Select(x => ctx.Document.GetTool(x))
            .Where(x => x is
            {
                CaptionData: { IsContinuation: false },
            } && x.CaptionData.Value.Type == type && predicate(x))
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
            if (elements[i].AssociatedHeading1?.HeadingData?.Number is {} headingNumber && allowHierarchical)
            {
                correctNumberSection = $"{headingNumber}.{underHeadingNumber}";
            }
            
            if (actualNumber == correctNumber)
            {
                if (hierarchicalNumbering == true)
                {
                    if (mixMessageId != null)
                        // TODO: figure out whether this is a content or formatting error.
                        ctx.AddMessage(new LintDiagnostic(mixMessageId, DiagnosticType.ContentError, new ParagraphDiagnosticContext(elements[i].Paragraph))
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
                    if (mixMessageId != null)
                        // TODO: figure out whether this is a content or formatting error.
                        ctx.AddMessage(new LintDiagnostic(mixMessageId, DiagnosticType.ContentError, new ParagraphDiagnosticContext(elements[i].Paragraph))
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

            // TODO: figure out whether this is a content or formatting error.
            ctx.AddMessage(new LintDiagnostic(messageId, DiagnosticType.ContentError, new ParagraphDiagnosticContext(elements[i].Paragraph))
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