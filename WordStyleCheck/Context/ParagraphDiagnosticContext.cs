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

    public void WriteCommentReference(string commentId)
    {
        if (Paragraphs[0].ParagraphProperties is { } props)
        {
            props.InsertAfterSelf(new CommentRangeStart()
            {
                Id = commentId
            });
        }
        else
        {
            Paragraphs[0].PrependChild(new CommentRangeStart()
            {
                Id = commentId
            });
        }

        Paragraphs[^1].Append(new CommentRangeEnd()
        {
            Id = commentId
        });
        Paragraphs[^1].Append(new Run(new CommentReference()
        {
            Id = commentId
        }));
    }
}