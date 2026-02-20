namespace WordStyleCheck;

public record LintMessage(string Message, bool AutoFixed, Context Context);