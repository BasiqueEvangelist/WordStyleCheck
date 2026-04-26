using System.IO.Hashing;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record RunDiagnosticContext(List<Run> Runs) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines
    {
        get
        {
            StringBuilder text = new StringBuilder();

            foreach (var r in Runs)
            {
                foreach (var t in r.Descendants<Text>())
                {
                    text.Append(t.Text);
                }
            }

            return [new("…", Utils.Truncate(text.ToString()), "…")];
        }
    }

    public RunDiagnosticContext(Run r) : this([r]) { }

    public void WriteToConsole()
    {
        // TODO: before and after
        
        StringBuilder text = new StringBuilder();
        
        foreach (var r in Runs)
        {
            foreach (var t in r.Descendants<Text>())
            {
                text.Append(t.Text);
            }
        }
        
        Console.Write(" |  ");

        Console.Write("…");
            
        if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
        Console.Write(Utils.Truncate(text.ToString()));
        if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            
        Console.Write("…");
            
        Console.WriteLine();
    }

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
    {
        Runs[0].InsertBeforeSelf(new CommentRangeStart()
        {
            Id = commentId
        });
        Runs[^1].InsertAfterSelf(new Run(new CommentReference()
        {
            Id = commentId
        }));
        Runs[^1].InsertAfterSelf(new CommentRangeEnd()
        {
            Id = commentId
        });
    }

    public IDiagnosticContext? TryMerge(IDiagnosticContext previous)
    {
        if (previous is not RunDiagnosticContext prevR) return null;
        
        OpenXmlElement? nextAfterPrev = prevR.Runs[^1];

        do
        {
            var orig = nextAfterPrev;
            nextAfterPrev = nextAfterPrev?.NextSibling();


            if (nextAfterPrev == null)
            {
                var p = Utils.AscendToAnscestor<Paragraph>(orig);

                if (p?.NextSibling() is Paragraph nextP)
                {
                    var children = Utils.DirectRunChildren(nextP);
                    nextAfterPrev = children.Count < 1 ? null : Utils.DirectRunChildren(nextP)[0];
                }
                else
                {
                    break;
                }
            }
            
            while (nextAfterPrev is ProofError) nextAfterPrev = nextAfterPrev.NextSibling();
        } while (nextAfterPrev != Runs[0] && nextAfterPrev is Run r && string.IsNullOrWhiteSpace(Utils.CollectText(r)));
            
        if (nextAfterPrev != Runs[0]) return null;

        return new RunDiagnosticContext([..prevR.Runs, ..Runs]);
    }
    
    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        hasher.Append(BitConverter.GetBytes(Runs.Count));
        foreach (var r in Runs)
        {
            HashUtils.HashElement(r, hasher);
        }
    }
}