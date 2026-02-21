using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Context;

public record ParagraphDiagnosticContext(List<Paragraph> Paragraphs) : IDiagnosticContext
{
    public ParagraphDiagnosticContext(Paragraph p) : this([p]) { }
    
    public void WriteToConsole()
    {
        foreach (var p in Paragraphs)
        {
            var text = Utils.CollectParagraphText(p, 25);

            Console.Write(" |  ");
            
            if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
            Console.Write(text.Text);
            if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            
            if (text.More) Console.Write("…");
            
            Console.WriteLine();
        }
    }
}