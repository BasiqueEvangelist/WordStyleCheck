namespace WordStyleCheck.Context;

public interface IDiagnosticContext
{
    List<DiagnosticContextLine> Lines { get; }

    void WriteToConsole();

    void WriteCommentReference(string commentId);

    IDiagnosticContext? TryMerge(IDiagnosticContext previous) => null;
}

public record struct DiagnosticContextLine(string Before, string Text, string After);