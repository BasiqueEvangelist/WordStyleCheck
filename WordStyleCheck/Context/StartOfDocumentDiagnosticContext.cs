using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record StartOfDocumentDiagnosticContext : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines => [];

    public void WriteToConsole()
    {
        
    }

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
    {
        var run = new Run(new CommentReference()
        {
            Id = commentId
        });

        var start = ctx.AllParagraphs.First();

        if (start.ParagraphProperties != null)
            start.ParagraphProperties.InsertAfterSelf(run);
        else
            start.PrependChild(run);
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        
    }
}