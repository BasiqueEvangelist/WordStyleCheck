using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintMessage(string Message, LintMessage.ExpectedActual? Values, bool AutoFixed, IDiagnosticContext Context)
{
    public record struct ExpectedActual(string Expected, string? Actual);
}