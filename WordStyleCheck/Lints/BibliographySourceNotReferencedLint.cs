using System.Text;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class BibliographySourceNotReferencedLint(Predicate<ParagraphPropertiesTool> isInBibliography) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["BibliographySourceNotReferenced"];

    public void Run(LintContext ctx)
    {
        HashSet<int> referenced = [];
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            string text = tool.Contents;

            int i = 0;

            int? ConsumeNumber()
            {
                StringBuilder numBuilder = new();

                while (i < text.Length && char.IsDigit(text[i]))
                {
                    numBuilder.Append(text[i]);
                    i += 1;
                }

                // TODO: span.
                if (int.TryParse(numBuilder.ToString(), out var res))
                    return res;
                else
                    return null;
            }
            
            void ConsumeReference()
            {
                i += 1;

                while (i < text.Length && text[i] != ']')
                {
                    if (char.IsDigit(text[i]))
                    {
                        int? first = ConsumeNumber();
                        if (first is null) return;

                        if (i < text.Length && text[i] is '-' or '–')
                        {
                            i += 1;

                            int? second = ConsumeNumber();
                            
                            if (second is not null)
                            {
                                for (int j = first.Value; j <= second; j++)
                                {
                                    referenced.Add(j);
                                }
                            }
                        }
                        else
                        {
                            referenced.Add(first.Value);
                        }
                    }
                    else
                    {
                        i += 1;
                    }
                }

                i += 1;
            }

            while (i < text.Length)
            {
                if (text[i] == '[')
                {
                    ConsumeReference();
                }
                else
                {
                    i += 1;
                }
            }
        }
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (!isInBibliography(tool)) continue;
            if (!(tool.OfNumbering is {} numbering)) continue;

            int index = numbering.Paragraphs.IndexOf(p) + 1;
            
            if (referenced.Contains(index)) continue;
            
            ctx.AddMessage(new LintDiagnostic("BibliographySourceNotReferenced", new ParagraphDiagnosticContext([p], true)));
        }
    }
}