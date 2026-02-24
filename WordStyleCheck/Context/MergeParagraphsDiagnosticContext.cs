using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Context;

public record MergeParagraphsDiagnosticContext(Paragraph First, Paragraph Second) : IDiagnosticContext
{
    public void WriteToConsole()
    {
        var fText = Utils.CollectParagraphText(First);
        var sText = Utils.CollectParagraphText(Second);
        
        Console.Write(" |  ");
        
        Console.Write("…");
        Console.Write(fText[Math.Max(fText.Length - 25, 0)..]);
        
        if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
        Console.Write("¶");
        if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            
        Console.Write(sText[..Math.Min(sText.Length, 25)]);
        Console.Write("…");
            
        Console.WriteLine();
    }
    
    public void WriteCommentReference(string commentId)
    {
        First.Append(new CommentRangeEnd()
        {
            Id = commentId
        });
        
        if (Second.ParagraphProperties is { } props)
        {
            props.InsertAfterSelf(new Run(new CommentReference()
            {
                Id = commentId
            }));
            props.InsertAfterSelf(new CommentRangeStart()
            {
                Id = commentId
            });
        }
        else
        {
            Second.PrependChild(new Run(new CommentReference()
            {
                Id = commentId
            }));
            Second.PrependChild(new CommentRangeStart()
            {
                Id = commentId
            });
        }
        
        
    }
}