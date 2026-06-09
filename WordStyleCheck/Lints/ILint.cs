namespace WordStyleCheck.Lints;

public interface ILint
{
    IReadOnlyList<string> EmittedDiagnostics { get; }
    
    void Run(ILintContext ctx);
}