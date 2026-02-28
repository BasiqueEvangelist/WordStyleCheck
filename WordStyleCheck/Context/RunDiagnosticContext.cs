using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Context;

public record RunDiagnosticContext(List<Run> Runs) : IDiagnosticContext
{
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

    public void WriteCommentReference(string commentId)
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
        
        var nextAfterPrev = prevR.Runs[^1].NextSibling();

        while (nextAfterPrev != Runs[0] &&
               (nextAfterPrev is Run r && string.IsNullOrWhiteSpace(Utils.CollectText(r))))
        {
            nextAfterPrev = nextAfterPrev.NextSibling();
        }
            
        if (nextAfterPrev != Runs[0]) return null;

        return new RunDiagnosticContext([..prevR.Runs, ..Runs]);
    }
}