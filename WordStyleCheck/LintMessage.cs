using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintMessage(string Message, bool AutoFixed, IDiagnosticContext Context);