using System.IO.Hashing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record SectionDiagnosticContext(SectionPropertiesTool Section) : IDiagnosticContext
{
    public List<DiagnosticContextLine> Lines
    {
        get
        {
            DiagnosticContextLine ParagraphToLine(Paragraph paragraph)
            {
                var text = Utils.CollectParagraphText(paragraph, 25);

                return new DiagnosticContextLine("", text.Text, text.More ? "…" : "");
            }

            if (Section.Paragraphs.Count <= 6)
                return Section.Paragraphs.Select(ParagraphToLine).ToList();
            else
                return Section.Paragraphs.Take(3).Select(ParagraphToLine)
                    .Append(new DiagnosticContextLine("…", "", ""))
                    .Concat(Section.Paragraphs.TakeLast(3).Select(ParagraphToLine))
                    .ToList();
        }
    }

    public void WriteToConsole()
    {
        // TODO: write proper context.
        foreach (var line in Lines)
        {
            Console.Write(" |  ");
            
            Console.Write(line.Before);
            if (!Console.IsOutputRedirected) Console.Write("\x1B[4m");
            Console.Write(line.Text);
            if (!Console.IsOutputRedirected) Console.Write("\x1B[0m");
            Console.Write(line.After);
            
            Console.WriteLine();
        }
    }

    public void WriteCommentReference(string commentId)
    {
        if (Section.Paragraphs[0].ParagraphProperties is { } props)
        {
            props.InsertAfterSelf(new CommentRangeStart()
            {
                Id = commentId
            });
        }
        else
        {
            Section.Paragraphs[0].PrependChild(new CommentRangeStart()
            {
                Id = commentId
            });
        }

        Section.Paragraphs[^1].Append(new CommentRangeEnd()
        {
            Id = commentId
        });
        Section.Paragraphs[^1].Append(new Run(new CommentReference()
        {
            Id = commentId
        }));
    }

    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        hasher.Append(BitConverter.GetBytes(Section.Context.AllSections.ToList().IndexOf(Section)));
    }
}