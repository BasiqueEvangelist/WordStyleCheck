using System.IO.Hashing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public interface IDiagnosticContext
{
    List<DiagnosticContextLine> Lines { get; }

    void WriteToConsole();

    void WriteCommentReference(string commentId, DocumentAnalysisContext ctx);

    IDiagnosticContext? TryMerge(IDiagnosticContext previous) => null;

    void Hash(NonCryptographicHashAlgorithm hasher);
}

public record struct DiagnosticContextLine(string Before, string Text, string After);