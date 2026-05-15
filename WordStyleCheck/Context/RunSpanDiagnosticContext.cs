using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public class RunSpanDiagnosticContext(RunSpan span) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines => [new("…", Utils.Truncate(span.ToString()), "…")];

    public void WriteToConsole()
    {
        // TODO: before and after
        
        Console.Write(" |  ");

        Console.Write("…");
            
        if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
        Console.Write(Utils.Truncate(span.ToString()));
        if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            
        Console.Write("…");
            
        Console.WriteLine();
    }

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
    {
        var runs = span.Isolate().ToList();
        
        runs[0].InsertBeforeSelf(new CommentRangeStart()
        {
            Id = commentId
        });
        runs[^1].InsertAfterSelf(new Run(new CommentReference()
        {
            Id = commentId
        }));
        runs[^1].InsertAfterSelf(new CommentRangeEnd()
        {
            Id = commentId
        });
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        span.Hash(hasher);
    }

    public bool AfterAll => true;
}