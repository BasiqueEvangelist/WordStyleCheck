namespace WordStyleCheck.Lints;

public interface ILint
{
    void Run(ILintContext ctx);
    
    IReadOnlyList<string> EmittedDiagnostics { get; }
}