using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public struct CaptionClassifierData
{
    public required CaptionType Type { get; init; }
    public required bool IsBelow { get; init; }
    public required bool IsContinuation { get; init; }
    public required string Number { get; init; }
    
    public required StringSpan TypeSpan { get; init; }
    public required StringSpan NumberSpan { get; init; }
    
    public static CaptionClassifierData? Classify(Paragraph p)
    {
        CaptionType type;
        bool isBelow;
        if (p.PreviousSibling() is Paragraph prev && prev.Descendants<Drawing>().Any())
        {
            type = CaptionType.Figure;
            isBelow = true;
        }
        else if (p.NextSibling() is Table)
        {
            type = CaptionType.Table;
            isBelow = false;
        }
        else if (p.NextSibling() is Paragraph next && next.Descendants<Drawing>().Any())
        {
            type = CaptionType.Figure;
            isBelow = false;
        } 
        else if (p.PreviousSibling() is Table)
        {
            type = CaptionType.Table;
            isBelow = true;
        }
        else
        {
            // Not attached to anything?
            return null;
        }

        string text = Utils.CollectParagraphText(p).Trim();
        
        int firstPartEnd;
        for (firstPartEnd = 0; firstPartEnd < text.Length; firstPartEnd++)
        {
            if (!(char.IsLetter(text[firstPartEnd]) || char.IsWhiteSpace(text[firstPartEnd]))) 
                break;
        }

        if (firstPartEnd >= text.Length) return null;

        string firstPart = text[..firstPartEnd].ToLowerInvariant();
        
        bool isContinuation = false;

        switch (type)
        {
            case CaptionType.Table when Algorithms.LevenshteinNeighbors(firstPart, "продолжение таблицы", 5):
                isContinuation = true;        
                break;
            case CaptionType.Figure when !Algorithms.LevenshteinNeighbors(firstPart, "рисунок", 2):
            case CaptionType.Table  when !Algorithms.LevenshteinNeighbors(firstPart, "таблица", 2):
                return null;
        }

        int secondPartEnd;
        for (secondPartEnd = firstPartEnd; secondPartEnd < text.Length; secondPartEnd++)
        {
            if (!char.IsLetter(text[secondPartEnd]) && !char.IsNumber(text[secondPartEnd]) && text[secondPartEnd] != '.') 
                break;
        }

        int secondPartStart = firstPartEnd;

        while (char.IsWhiteSpace(text[secondPartStart]))
        {
            secondPartStart += 1;
        }

        return new CaptionClassifierData
        {
            Type = type,
            IsBelow = isBelow,
            IsContinuation = isContinuation,
            Number = text.Substring(secondPartStart, secondPartEnd - secondPartStart).Trim(),
            
            TypeSpan = new StringSpan(0, firstPartEnd),
            NumberSpan = new StringSpan(secondPartStart, secondPartEnd)
        };
    }
}

public enum CaptionType
{
    Figure,
    Table,
}

public record struct StringSpan(int BeginInclusive, int EndExclusive);