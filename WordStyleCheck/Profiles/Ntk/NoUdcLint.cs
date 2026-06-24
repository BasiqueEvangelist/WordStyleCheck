using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class NoUdcLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["NoUdc"];
    
    public void Run(ILintContext ctx)
    {
        bool any = ctx.Document.AllParagraphs.Select(ctx.Document.GetTool).Any(x => x.GetFeature(NtkParagraphData.Key)!.IsUniversalDecimalClassifier);

        if (any) return;

        ctx.AddMessage(new LintDiagnostic(
            "NoUdc",
            DiagnosticType.ContentError,
            new StartOfDocumentDiagnosticContext()));
            
        if (ctx.AutomaticallyFix)
        {
            Paragraph p = new Paragraph();
            Run r = new Run(new Text("УДК 000.000"));

            p.AppendChild(r);

            r.RunProperties = new RunProperties();
            r.RunProperties.Highlight = new Highlight();
            r.RunProperties.Highlight.Val = HighlightColorValues.Red;
            r.RunProperties.RunFonts = new RunFonts();
            r.RunProperties.RunFonts.Ascii = "something that the text font lint will correct";

            ctx.Document.Document.MainDocumentPart!.Document!.Body!.PrependChild(p);

            var tool = ctx.Document.GetTool(p);
            tool.SetFeature(NtkParagraphData.Key, new NtkParagraphData(tool));
            tool.GetFeature(NtkParagraphData.Key)!.IsUniversalDecimalClassifier = true;
            
            ctx.Document.ReloadParagraphs();
        }
    }
}