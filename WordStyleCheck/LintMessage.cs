using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintMessage(
    string Message,
    IDiagnosticContext Context,
    Action? AutoFix = null,
    LintMessage.ExpectedActual? Values = null
)
{
    public record struct ExpectedActual(string Expected, string? Actual);
}