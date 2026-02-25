using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintMessage(
    string Id,
    IDiagnosticContext Context,
    Dictionary<string, string>? Parameters = null,
    Action? AutoFix = null
)
{
    public record struct ExpectedActual(string Expected, string? Actual);
}