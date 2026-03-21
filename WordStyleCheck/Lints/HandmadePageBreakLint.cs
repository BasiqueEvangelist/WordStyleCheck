using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints
{
    public class HandmadePageBreakLint : ILint
    {
        public IReadOnlyList<string> EmittedDiagnostics { get; } = ["HandmadePageBreak"];

        public void Run(LintContext ctx)
        {
            int emptyParagraphsCount = 0;

            var paragraphs = ctx.Document.AllParagraphs.ToList();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                bool isEmpty = ctx.Document.GetTool(paragraphs[i]) is { IsEmptyOrWhitespace: true, Class: not ParagraphClass.CodeListing};

                if (!isEmpty || emptyParagraphsCount > 0 && paragraphs[i - 1].NextSibling() != paragraphs[i])
                {
                    if (emptyParagraphsCount > 3)
                    {
                        var chosen = paragraphs[(i - emptyParagraphsCount)..i].ToList();

                        ctx.AddMessage(new LintMessage("HandmadePageBreak",
                            new ParagraphDiagnosticContext(chosen, true)));
                    }

                    emptyParagraphsCount = 0;
                }
                
                if (isEmpty)
                {
                    emptyParagraphsCount++;
                }
            }
        }
    }
}
