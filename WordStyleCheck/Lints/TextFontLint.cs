using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class TextFontLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool pTool = ctx.Document.GetTool(p);

            if (pTool.IsEmptyOrDrawing) continue;
            
            // We don't force a specific font for code listings, since nobody knows what font they should actually use.
            if (pTool.Class == ParagraphClass.CodeListing) continue;

            // Drawings can have their own stuff. We don't really care.
            if (pTool.Class == ParagraphClass.InsideDrawing) continue;

            foreach (var r in p.Descendants<Run>())
            {
                if (string.IsNullOrWhiteSpace(Utils.CollectText(r))) continue;
                
                RunPropertiesTool tool = ctx.Document.GetTool(r);

                if (tool.AsciiFont != "Times New Roman")
                {
                    ctx.AddMessage(new LintMessage("TextFontIncorrect", new RunDiagnosticContext(r))
                    {
                        Parameters = new() {
                            ["Expected"] = "Times New Roman",
                            ["Actual"] = tool.AsciiFont ?? "<?>"
                        },
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