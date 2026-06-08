using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ZapFootnotesEndnotesLint(Predicate<ParagraphPropertiesTool> isInBibliography) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["FootnoteForbidden", "FootnoteForbiddenFound"];
    
    public void Run(ILintContext ctx)
    {
        Dictionary<int, int> footnoteReplacements = [];
        Dictionary<int, int> endnoteReplacements = [];

        List<ParagraphPropertiesTool> biblioEntries = ctx.Document.AllParagraphs
            .Select(ctx.Document.GetTool)
            .Where(isInBibliography.Invoke)
            .ToList();

        void MapNotes(IEnumerable<FootnoteEndnoteType> nodes, Dictionary<int, int> replacements)
        {
            foreach (var note in nodes)
            {
                string text = Utils.ToPlainText(note.ChildElements.ToList()).Trim();

                var candidates = biblioEntries
                    .Index()
                    .Select(x => new { x.Index, Distance = StringEx.LevenshteinDistance(text, x.Item.Contents.Trim()) })
                    .Where(x => x.Distance < text.Length / 5)
                    .ToList();
                
                int? index = candidates
                    .MinBy(x => x.Distance)
                    ?.Index;

                if (index == null) continue;

                replacements[(int)note.Id!.Value] = index.Value + 1; 
            
                if (ctx.AutomaticallyFix)
                {
                    ctx.MarkAutoFixed();
                    note.Remove();
                }
            }
        }
        
        MapNotes(ctx.Document.Footnotes.Values, footnoteReplacements);
        MapNotes(ctx.Document.Endnotes.Values, endnoteReplacements);

        foreach (var p in ctx.Document.AllParagraphs)
        {
            foreach (var run in Utils.DirectRunChildren(p))
            {
                foreach (var c in run.ChildElements.ToList())
                {
                    if (c is not FootnoteEndnoteReferenceType f) continue;
                    
                    Dictionary<int, int> replacements;

                    if (f is FootnoteReference)
                        replacements = footnoteReplacements;
                    else if (f is EndnoteReference)
                        replacements = endnoteReplacements;
                    else
                        throw new NotImplementedException();
                        
                    if (replacements.TryGetValue((int)f.Id!.Value, out var repl))
                    {
                        if (!ctx.AutomaticallyFix)
                        {
                            // TODO: highlight footnote reference.
                            ctx.AddMessage(new LintDiagnostic("FootnoteForbiddenFound", DiagnosticType.FormattingError, new RunDiagnosticContext(run))
                            {
                                Parameters = new()
                                {
                                    ["ReplacementNumber"] = repl.ToString()
                                }
                            });
                        }
                        else
                        {
                            ctx.MarkAutoFixed();

                            Text t = new Text()
                            {
                                Space = SpaceProcessingModeValues.Preserve,
                                Text = " [" + repl + "]"
                            };

                            f.InsertAfterSelf(t);
                            f.Remove();
                            
                            var rTool = ctx.Document.GetTool(run);

                            if (rTool.VerticalAlignment != VerticalPositionValues.Baseline)
                            {
                                run.RunProperties ??= new RunProperties();
                                if (ctx.GenerateRevisions) Utils.SnapshotRunProperties(run.RunProperties);

                                run.RunProperties.VerticalTextAlignment ??= new VerticalTextAlignment();
                                run.RunProperties.VerticalTextAlignment.Val = VerticalPositionValues.Baseline;
                            }
                        }
                    }
                    else
                    {
                        // TODO: highlight footnote reference.
                        ctx.AddMessage(new LintDiagnostic("FootnoteForbidden", DiagnosticType.FormattingError, new RunDiagnosticContext(run)));
                    }
                }
            }
        }
    }
}