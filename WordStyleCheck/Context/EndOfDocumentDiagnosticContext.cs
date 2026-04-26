using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record EndOfDocumentDiagnosticContext : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines => [];

    public void WriteToConsole()
    {
        
    }

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
    {
        ctx.AllParagraphs.Last().Append(new Run(new CommentReference()
        {
            Id = commentId
        }));
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        
    }
}