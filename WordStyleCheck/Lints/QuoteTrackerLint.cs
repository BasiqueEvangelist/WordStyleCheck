using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class QuoteTrackerLint(Predicate<ParagraphPropertiesTool> predicate) : ILint
{
    private const string Starts = "«“";
    private const string Ends   = "»”";

    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["IncorrectQuoteChoice", "UnpairedStartQuote", "UnpairedEndQuote", "WrongQuotePair", "QuotesTooDeep"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;
            if (!predicate(tool)) continue;

            List<(int, char)> quoteStack = [];

            RunAssociatedText rat = RunAssociatedText.FromParagraph(tool);

            for (int i = 0; i < rat.Text.Length; i++)
            {
                if (Utils.IsMonospaceFont(rat.GetRunAt(i).AsciiFont ?? "<?>")) continue;

                int startIdx = Starts.IndexOf(rat.Text[i]);

                if (startIdx != -1)
                {
                    quoteStack.Add((i, Ends[startIdx]));
                }

                if (Ends.Contains(rat.Text[i]))
                {
                    if (quoteStack.Count == 0)
                    {
                        var span = rat.GetSpan(i, 1);

                        ctx.AddMessage(new LintDiagnostic("UnpairedEndQuote", DiagnosticType.ContentError,
                            new RunSpanDiagnosticContext(span)));
                    }
                    else
                    {
                        var (beginningIdx, properLast) = quoteStack[^1];
                        quoteStack.RemoveAt(quoteStack.Count - 1);
                        var span = rat.GetSpan(beginningIdx, i - beginningIdx + 1);

                        if (quoteStack.Count >= 2)
                        {
                            ctx.AddMessage(new LintDiagnostic("QuotesTooDeep", DiagnosticType.ContentError,
                                new RunSpanDiagnosticContext(span)));
                        }
                        else if (properLast != Ends[quoteStack.Count])
                        {
                            if (!ctx.AutomaticallyFix)
                            {
                                ctx.AddMessage(new LintDiagnostic("IncorrectQuoteChoice", DiagnosticType.FormattingError,
                                    new RunSpanDiagnosticContext(span))
                                {
                                    Parameters = new()
                                    {
                                        ["ExpectedLeft"] = Starts[quoteStack.Count].ToString(),
                                        ["ActualLeft"] = rat.Text[beginningIdx].ToString(),
                                        
                                        ["ExpectedRight"] = Ends[quoteStack.Count].ToString(),
                                        ["ActualRight"] = properLast.ToString()
                                    }
                                });
                            }
                            else
                            {
                                ctx.MarkAutoFixed();

                                // TODO: make this more efficient
                                var lSpan = rat.GetSpan(beginningIdx, 1);
                                lSpan.Replace(Starts[quoteStack.Count].ToString());

                                tool.ReloadContents();
                                rat = RunAssociatedText.FromParagraph(tool);
                                
                                var rSpan = rat.GetSpan(i, 1);
                                rSpan.Replace(Ends[quoteStack.Count].ToString());
                                
                                tool.ReloadContents();
                                rat = RunAssociatedText.FromParagraph(tool);
                            }
                        }
                        else if (properLast != rat.Text[i])
                        {
                            if (!ctx.AutomaticallyFix)
                            {
                                ctx.AddMessage(new LintDiagnostic("WrongQuotePair", DiagnosticType.FormattingError,
                                    new RunSpanDiagnosticContext(span))
                                {
                                    Parameters = new()
                                    {
                                        ["Expected"] = properLast.ToString(),
                                        ["Actual"] = rat.Text[i].ToString()
                                    }
                                });
                            }
                            else
                            {
                                ctx.MarkAutoFixed();

                                // TODO: make this more efficient
                                var oSpan = rat.GetSpan(i, 1);
                                oSpan.Replace(properLast.ToString());
                                
                                tool.ReloadContents();
                                rat = RunAssociatedText.FromParagraph(tool);
                            }
                        }
                    }
                }
            }

            while (quoteStack.Count > 0)
            {
                var (beginningIdx, _) = quoteStack[^1];
                quoteStack.RemoveAt(quoteStack.Count - 1);
                
                ctx.AddMessage(new LintDiagnostic("UnpairedStartQuote", DiagnosticType.ContentError,
                    new RunSpanDiagnosticContext(rat.GetSpan(beginningIdx, 1))));
            }
        }
    }
}