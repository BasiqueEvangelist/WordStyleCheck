using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class BodyTextFontLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);
            
            if (pTool.Class != ParagraphClass.BodyText)
            {
                continue;
            }

            foreach (var r in p.Descendants<Run>())
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.AsciiFont != "Times New Roman")
                {
                    ctx.AddMessage(new LintMessage("Run doesn't have needed font", new RunDiagnosticContext(r))
                    {
                        Values = new("Times New Roman", tool.AsciiFont),
                        AutoFix = () =>
                        {
                            if (r.RunProperties == null) r.RunProperties = new RunProperties();
                            if (r.RunProperties.RunFonts == null) r.RunProperties.RunFonts = new RunFonts();

                            r.RunProperties.RunFonts.Ascii = "Times New Roman";
                            r.RunProperties.RunFonts.HighAnsi = "Times New Roman";
                            r.RunProperties.RunFonts.ComplexScript = "Times New Roman";
                        }
                    });
                }
            }
        }
    }
}