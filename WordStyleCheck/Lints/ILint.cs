namespace WordStyleCheck.Lints;

public interface ILint
{
    void Run(LintContext ctx);
    
    IReadOnlyList<string> EmittedDiagnostics { get; }
}