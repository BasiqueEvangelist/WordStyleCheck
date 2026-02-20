using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public enum ContextType
{
    Paragraph,
    Run
}

public record Context(ContextType Type, List<object> ContextObjects, string Before, string Text, string After)
{
    public static Context FromParagraph(Paragraph p)
    {
        var context = Utils.CollectParagraphText(p, 20);

        return new Context(ContextType.Paragraph, [p], "", context.Text[..Math.Min(20, context.Text.Length)], context.More ? "…" : "");
    }
    
    public static Context FromRun(Run r)
    {
        StringBuilder runText = new();
        foreach (var text in r.Descendants<Text>())
        {
            runText.Append(text.Text);
        }

        return new Context(ContextType.Run, [r], "", runText.ToString(), "");
    }
    
    public void WriteToConsole()
    {
        Console.Write(" |  ");
        Console.Write(Before);
        if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
        Console.Write(Text);
        if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
        Console.Write(After);
        Console.WriteLine();
    }
}
