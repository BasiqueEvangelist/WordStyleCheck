using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints
{
    public class IncorrectOutlineLevelLint(Predicate<ParagraphPropertiesTool> predicate, Func<ParagraphPropertiesTool, int?> outlineLevel, string messageId) : ILint
    {
        public void Run(LintContext ctx)
        {
            foreach (var p in ctx.Document.AllParagraphs)
            {
                ParagraphPropertiesTool tool = ctx.Document.GetTool(p);

                if (tool.IsEmptyOrDrawing) continue;
                if (!predicate(tool)) continue;

                if (tool.OutlineLevel != outlineLevel(tool))
                {
                    ctx.AddMessage(new LintMessage(messageId, new ParagraphDiagnosticContext(p)));
                }
            }
        }
    }
}
