using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class CorrectStructuralElementHeaderLint : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (tool.StructuralElementHeader == null) continue;

            var text = Utils.CollectParagraphText(p).Trim();
            var proper = StructuralElementHeaderClassifier.GetProperName(tool.StructuralElementHeader.Value);

            if (text != proper)
            {
                ctx.AddMessage(new LintMessage("StructuralElementHeaderContentsIncorrect", new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = proper,
                        ["Actual"] = text
                    },
                    // TODO: re-add autofix once it actually works properly
                    // AutoFix = () =>
                    // {
                    //     // TODO: add proper support for generate-revisions.
                    //
                    //     foreach (var child in p.ChildElements)
                    //     {
                    //         if (child is ParagraphProperties) continue;
                    //         
                    //         child.Remove();
                    //     }
                    //
                    //     p.Append(new Run(new Text(proper)));
                    // }
                });
            }
        }
    }
}