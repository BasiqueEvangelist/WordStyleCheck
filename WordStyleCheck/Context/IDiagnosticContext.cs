namespace WordStyleCheck.Context;

public interface IDiagnosticContext
{
    void WriteToConsole();

    void WriteCommentReference(string commentId);
}