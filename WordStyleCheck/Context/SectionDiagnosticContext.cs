using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Context;

public record SectionDiagnosticContext(SectionPropertiesTool Section) : IDiagnosticContext
{
    public void WriteToConsole()
    {
        // TODO: write proper context.
        foreach (var p in Section.Paragraphs)
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
}