using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Gost7_32;

public class CorrectStructuralElementHeaderLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["StructuralElementHeaderContentsIncorrect"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var data = ctx.Document.GetTool(p).GetFeature(GostParagraphData.Key)!;

            if (data.StructuralElementHeader == null) continue;
            if (data.StructuralElementHeader == GostStructuralElement.Appendix) continue;

            var text = data.Inner.Contents.Trim();
            var proper = GostStructuralElementClassifier.GetProperName(data.StructuralElementHeader.Value);

            if (text != proper)
            {
                ctx.AddMessage(new LintDiagnostic("StructuralElementHeaderContentsIncorrect", DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p))
                {
                    Parameters = new()
                    {
                        ["Expected"] = proper,
                        ["Actual"] = text
                    },
                    AutoFix = () =>
                    {
                        // TODO: add proper support for generate-revisions.
                    
                        foreach (var child in p.ChildElements.ToList())
                        {
                            if (child is ParagraphProperties) continue;
                            
                            child.Remove();
                        }
                    
                        p.Append(new Run(new Text(proper)));
                    }
                });
            }
        }
    }
}