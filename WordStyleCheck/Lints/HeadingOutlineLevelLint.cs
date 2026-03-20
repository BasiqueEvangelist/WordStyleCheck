using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HeadingOutlineLevelLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["NonHeadingWithOutlineLevel", "HeadingWithoutOutlineLevel"];

    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (tool.OfStructuralElement == null) continue;

            if (tool.OutlineLevel != null && tool.HeadingData == null)
            {
                ctx.AddMessage(new LintMessage("NonHeadingWithOutlineLevel", new ParagraphDiagnosticContext(p))
                {
                    AutoFix = () =>
                    {
                        if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();

                        if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                        p.ParagraphProperties.OutlineLevel?.Remove();

                        p.ParagraphProperties.OutlineLevel = new OutlineLevel()
                        {
                            Val = 9
                        };
                    }
                });
            }

            if (tool.HeadingData != null && tool.OutlineLevel == null)
            {
                ctx.AddMessage(new LintMessage("HeadingWithoutOutlineLevel", new ParagraphDiagnosticContext(p)));
            }
        }
    }
}
