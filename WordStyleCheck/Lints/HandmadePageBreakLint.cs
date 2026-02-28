using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints
{
    public class HandmadePageBreakLint : ILint
    {
        public void Run(LintContext ctx)
        {
            int emptyParagraphsCount = 0;

            var paragraphs = ctx.Document.AllParagraphs.ToList();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(paragraphs[i])))
                {
                    emptyParagraphsCount++;
                }
                else
                {
                    if (emptyParagraphsCount > 3)
                    {
                        var chosen = paragraphs[(i - emptyParagraphsCount)..i].ToList();

                        ctx.AddMessage(new LintMessage("HandmadePageBreak", new ParagraphDiagnosticContext(chosen, true)));
                    }

                    emptyParagraphsCount = 0;
                }
            }
        }
    }
}
