using DocumentFormat.OpenXml.Packaging;

namespace WordStyleCheck.Lints;

public interface ILint
{
    void Run(LintContext ctx);
}