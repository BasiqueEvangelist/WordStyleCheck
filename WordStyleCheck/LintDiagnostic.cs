using System.IO.Hashing;
using System.Text;
using WordStyleCheck.Context;

namespace WordStyleCheck;

public record LintDiagnostic(
    string Id,
    DiagnosticType Type,
    IDiagnosticContext Context,
    Dictionary<string, string>? Parameters = null,
    Action? AutoFix = null
)
{
    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        hasher.Append(Encoding.UTF8.GetBytes(Id));
        hasher.Append(BitConverter.GetBytes((int) Type));
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

    public string GetHash()
    {
        XxHash128 hasher = new XxHash128();
        Hash(hasher);
        return Convert.ToHexString(hasher.GetCurrentHash());
    }
}

public enum DiagnosticType
{
    /// <summary>
    /// Lint or classifier failed. Guru meditation.
    /// </summary>
    Fatal,
    
    /// <summary>
    /// Invalid .docx file.
    /// </summary>
    CouldNotOpen,
    
    /// <summary>
    /// Invalid document structure.
    /// </summary>
    CouldNotParse,
    
    /// <summary>
    /// Basic formatting error, that can (usually) be easily automatically fixed. 
    /// </summary>
    FormattingError,
    
    /// <summary>
    /// Error in the content of the document. Cannot be automatically fixed.
    /// </summary>
    ContentError,
    
    /// <summary>
    /// Error in document content that can result from the document being incomplete.
    /// </summary>
    IncompleteContentError,
}