using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record MergeParagraphsDiagnosticContext(Paragraph First, Paragraph Second) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines
    {
        get
        {
            var fText = Utils.CollectParagraphText(First);
            var sText = Utils.CollectParagraphText(Second);

            return [new(fText[Math.Max(fText.Length - 25, 0)..], "¶", sText[..Math.Min(sText.Length, 25)])];
        }
    }

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
    
    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
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

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        HashUtils.HashElement(First, hasher);
        HashUtils.HashElement(Second, hasher);
    }
}