using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints
{
    public class NotEnoughSourcesLint(int neededCount, string messageId, string noBibliographyMessageId) : ILint
    {
        public void Run(LintContext ctx)
        {
            var numberings = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Where(x => x is { OfStructuralElement: Analysis.StructuralElement.Bibliography, OfNumbering: not null }).Select(x => x.OfNumbering!).Distinct().ToList();

            if (numberings.Count == 0)
            {
                ctx.AddMessage(new LintMessage(noBibliographyMessageId, new EndOfDocumentDiagnosticContext(ctx.Document.AllParagraphs.Last())));
                return;
            }
            
            int totalSources = numberings.Sum(x => x.Paragraphs.Count);

            if (totalSources < neededCount)
            {
                ctx.AddMessage(new LintMessage(messageId, new ParagraphDiagnosticContext(numberings.SelectMany(x => x.Paragraphs).ToList()))
                {
                    Parameters = new()
                    {
                        ["Expected"] = neededCount.ToString(),
                        ["Actual"] = totalSources.ToString()
                    }
                });
            }
        }
    }
}
