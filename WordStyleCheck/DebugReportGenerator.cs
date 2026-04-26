using System.IO.Hashing;
using System.Text;

namespace WordStyleCheck;

public class DebugReportGenerator(TextWriter writer)
{
    public void WriteHeader(string additionalText)
    {
        // TODO: write in actual version.
        writer.WriteLine("WordStyleCheck 0.0.1");
        writer.WriteLine(additionalText);
        writer.WriteLine();
    }

    public void WriteDiagnostic(LintDiagnostic diagnostic, XmlTranslationsFile translations)
    {
        string code = diagnostic.GetHash();
        
        writer.WriteLine($"-------- {code} --------");
        
        StringBuilder dumped = new StringBuilder(diagnostic.Id);

        if (diagnostic.Parameters is { Count: > 0 })
        {
            dumped.Append(" {");
            dumped.AppendJoin(", ", diagnostic.Parameters.Select(x => $"{x.Key} = '{x.Value}'"));
            dumped.Append("}");
        }

        writer.WriteLine(dumped.ToString());
        
        writer.WriteLine();
        
        writer.WriteLine(Utils.ToPlainText(translations.Translate(diagnostic.Id, diagnostic.Parameters ?? new(), null)));
        
        writer.WriteLine();

        foreach (var line in diagnostic.Context.Lines)
        {
            writer.WriteLine("Context:");
            writer.WriteLine("  Before: " + line.Before);
            writer.WriteLine("    Text: " + line.Text);
            writer.WriteLine("   After: " + line.After);
        }
        
        writer.WriteLine();
    }
}