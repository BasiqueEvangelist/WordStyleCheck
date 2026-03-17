using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Context;

public record StartOfDocumentDiagnosticContext(Paragraph StartParagraph) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines => [];

    public void WriteToConsole()
    {
        
    }

    public void WriteCommentReference(string commentId)
    {
        var run = new Run(new CommentReference()
        {
            Id = commentId
        });

        if (StartParagraph.ParagraphProperties != null)
            StartParagraph.ParagraphProperties.InsertAfterSelf(run);
        else
            StartParagraph.PrependChild(run);
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        
    }
}