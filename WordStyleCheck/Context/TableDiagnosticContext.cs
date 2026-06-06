using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record TableDiagnosticContext(Table Table) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines { get; } = [new DiagnosticContextLine()
    {
        Text = Utils.CollectParagraphText(Table.Descendants<Paragraph>().First(), 50).Text
    }];
    public void WriteToConsole()
    {
        var text = Utils.CollectParagraphText(Table.Descendants<Paragraph>().First(), 25);

        Console.Write(" |  ");
            
        if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
        Console.Write(text.Text);
        if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            
        if (text.More) Console.Write("…");
            
        Console.WriteLine();
    }

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
    {
        Table.PrependChild(new CommentRangeStart()
        {
            Id = commentId
        });
        Table.AppendChild(new CommentRangeEnd()
        {
            Id = commentId
        });

        Table.Descendants<Run>().LastOrDefault()?.AppendChild(new CommentReference()
        {
            Id = commentId
        });
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        HashUtils.HashElement(Table, hasher);
    }
}