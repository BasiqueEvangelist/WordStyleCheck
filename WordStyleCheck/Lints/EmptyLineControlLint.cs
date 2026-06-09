using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class EmptyLineControlLint(List<EmptyLineControlLint.Rule> beforeRules, List<EmptyLineControlLint.Rule> afterRules, string unneededId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = beforeRules.Select(x => x.MissingId)
        .Concat(afterRules.Select(x => x.MissingId)).Append(unneededId).ToList();
    
    public void Run(ILintContext ctx)
    {
        HashSet<ParagraphPropertiesTool> allowedLines = [];
        
        foreach (var e in ctx.Document.AllBlockLevel)
        {
            IBlockLevelPropertiesTool tool;
            IDiagnosticContext diagCtx;

            if (e is Paragraph p)
            {
                tool = ctx.Document.GetTool(p);
                diagCtx = new ParagraphDiagnosticContext(p);
            }
            else if (e is Table t)
            {
                tool = ctx.Document.GetTool(t);
                diagCtx = new TableDiagnosticContext(t);
            }
            else
                throw new NotImplementedException();


            void CheckRules(List<Rule> rules, bool before)
            {
                foreach (var rule in rules)
                {
                    if (!rule.LinePredicate(tool)) continue;

                    OpenXmlElement? sibling = before ? e.PreviousSibling() : e.NextSibling();
                    
                    if (sibling is Paragraph emptyP && ctx.Document.GetTool(emptyP) is {IsEmptyOrWhitespace: true} emptyTool)
                    {
                        allowedLines.Add(emptyTool);
                        
                        foreach (var lint in rule.Lints)
                        {
                            lint.Run(ctx, emptyTool);
                        }
                    }
                    else
                    {
                        if (!ctx.AutomaticallyFix)
                        {
                            ctx.AddMessage(new LintDiagnostic(rule.MissingId, DiagnosticType.FormattingError, diagCtx));
                        }
                        else
                        {
                            ctx.MarkAutoFixed();
                            
                            Paragraph newEmpty = new Paragraph();

                            if (before)
                                e.InsertBeforeSelf(newEmpty);
                            else
                                e.InsertAfterSelf(newEmpty);
                        
                            ParagraphPropertiesTool newEmptyTool = ctx.Document.GetTool(newEmpty);

                            foreach (var lint in rule.Lints)
                            {
                                lint.Run(ctx, newEmptyTool);
                            }
                        }
                    }
                }
            }
            
            CheckRules(beforeRules, true);
            CheckRules(afterRules, false);
        }

        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!tool.IsEmptyOrWhitespace) continue;
            if (allowedLines.Contains(tool)) continue;

            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic(unneededId, DiagnosticType.FormattingError, new ParagraphDiagnosticContext(p)));
            }
            else
            {
                ctx.MarkAutoFixed();
                p.Remove();
            }
        }
    }
    
    public record struct Rule(Predicate<IBlockLevelPropertiesTool> LinePredicate, string MissingId, List<ParagraphLint> Lints);
}