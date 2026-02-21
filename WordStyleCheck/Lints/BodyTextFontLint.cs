using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Lints;

public class BodyTextFontLint : ILint
{
    public void Run(LintContext ctx)
    {
        var body = ctx.Document.MainDocumentPart?.Document?.Body;

        if (body == null) return;

        foreach (var p in body.Descendants<Paragraph>())
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool pTool = new(ctx.Document, p);
            
            if (pTool.OutlineLevel != null)
            {
                // Probably not body text.
                Console.WriteLine(pTool.Style);
                continue;
            }

            foreach (var r in p.Descendants<Run>())
            {
                RunPropertiesTool tool = new(ctx.Document, r);

                if (tool.AsciiFont != "Times New Roman")
                {
                    ctx.AddMessage(new LintMessage($"Run doesn't have needed font (expected 'Times New Roman', was '{tool.AsciiFont}')", ctx.AutofixEnabled, Context.FromRun(r)));

                    if (ctx.AutofixEnabled)
                    {
                        ctx.MarkDocumentChanged();

                        if (r.RunProperties == null) r.RunProperties = new RunProperties();
                        if (r.RunProperties.RunFonts == null) r.RunProperties.RunFonts = new RunFonts();

                        r.RunProperties.RunFonts.Ascii = "Times New Roman";
                        r.RunProperties.RunFonts.HighAnsi = "Times New Roman";
                        r.RunProperties.RunFonts.ComplexScript = "Times New Roman";
                    }
                }
            }
        }
    }
}