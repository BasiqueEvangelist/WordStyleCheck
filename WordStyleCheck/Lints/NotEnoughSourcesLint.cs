using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints
{
    public class NotEnoughSourcesLint(int neededCount, string messageId) : ILint
    {
        public void Run(LintContext ctx)
        {
            var numberings = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Where(x => x.OfStructuralElement == Analysis.StructuralElement.Bibliography && x.OfNumbering != null).Select(x => x.OfNumbering).Distinct().ToList();

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
