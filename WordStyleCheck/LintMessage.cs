using System.IO.Hashing;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintMessage(
    string Id,
    IDiagnosticContext Context,
    Dictionary<string, string>? Parameters = null,
    Action? AutoFix = null
)
{
    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        hasher.Append(Encoding.UTF8.GetBytes(Id));
        Context.Hash(hasher);

        if (Parameters != null)
        {
            foreach (var entry in Parameters)
            {
                hasher.Append(Encoding.UTF8.GetBytes(entry.Key));
                hasher.Append(Encoding.UTF8.GetBytes(entry.Value));
            }
        }
        
        hasher.Append(BitConverter.GetBytes(AutoFix != null));
    }
}