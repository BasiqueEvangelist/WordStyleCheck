using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public static class StripOldComments
{
    public static void Run(WordprocessingDocument document)
    {
        var comments = document.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
        if (comments == null) return;

        HashSet<string> idsToRemove = [];
            
        foreach (var comment in comments.ChildElements.OfType<Comment>().ToList())
        {
            if (comment.Author?.Value == "WordStyleCheck" && comment.Initials?.Value == "WSC")
            {
                comment.Remove();
                idsToRemove.Add(comment.Id!.Value!);
            }
        }

        foreach (var toRemove in document.MainDocumentPart!.Document!.Descendants()
                     .Where(x => x is CommentReference or CommentRangeStart or CommentRangeEnd).ToList())
        {
            if ((toRemove is CommentReference { Id.Value: { } idRef } && idsToRemove.Contains(idRef))
             || (toRemove is CommentRangeStart { Id.Value: { } idStart } && idsToRemove.Contains(idStart))
             || (toRemove is CommentRangeEnd { Id.Value: { } idEnd } && idsToRemove.Contains(idEnd)))
            {
                toRemove.Remove();
            }
        }
    }
}