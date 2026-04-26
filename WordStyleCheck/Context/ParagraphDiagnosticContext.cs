using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record ParagraphDiagnosticContext(List<Paragraph> Paragraphs, bool DisableMerging = false) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines => Paragraphs.Select(x =>
    {
        var text = Utils.CollectParagraphText(x, 25);

        return new DiagnosticContextLine("", text.Text, text.More ? "…" : "");
    }).ToList();

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

    public void WriteCommentReference(string commentId, DocumentAnalysisContext ctx)
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

    public IDiagnosticContext? TryMerge(IDiagnosticContext previous)
    {
        if (previous is not ParagraphDiagnosticContext prevP) return null;

        if (prevP.DisableMerging || DisableMerging) return null;
        
        if (prevP.Paragraphs[^1].NextSibling() != Paragraphs[0]) return null;

        return new ParagraphDiagnosticContext([..prevP.Paragraphs, ..Paragraphs]);
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        hasher.Append(BitConverter.GetBytes(Paragraphs.Count));
        foreach (var p in Paragraphs)
        {
            HashUtils.HashElement(p, hasher);
        }
        
        hasher.Append(BitConverter.GetBytes(DisableMerging));
    }
}