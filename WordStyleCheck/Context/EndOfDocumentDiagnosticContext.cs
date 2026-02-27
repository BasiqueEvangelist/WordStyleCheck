using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Context;

public record EndOfDocumentDiagnosticContext(Paragraph EndParagraph) : IDiagnosticContext
{
    public void WriteToConsole()
    {
        
    }

    public void WriteCommentReference(string commentId)
    {
        EndParagraph.Append(new Run(new CommentReference()
        {
            Id = commentId
        }));
    }
}