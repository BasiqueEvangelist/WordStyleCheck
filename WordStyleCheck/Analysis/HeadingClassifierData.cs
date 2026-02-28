namespace WordStyleCheck.Analysis;

public class HeadingClassifierData
{
    public required string Number { get; init; }
    public required string Title { get; init; }
    
    public static HeadingClassifierData? Classify(ParagraphPropertiesTool p)
    {
        if (p.OfNumbering != null) return null;
        if (p.IsEmptyOrDrawing) return null;
        if (p.Class is not (ParagraphClass.BodyText or ParagraphClass.Heading)) return null;
        
        string text = Utils.CollectParagraphText(p.Paragraph).Trim();

        if (text.Length < 1) return null;

        int numEnd = 0;

        while (numEnd < text.Length && (char.IsDigit(text[numEnd]) || text[numEnd] == '.'))
            numEnd += 1;

        if (numEnd == text.Length) return null;
        
        string number = text[..numEnd].TrimEnd('.');

        string[] numberSplit = number.Split('.');

        if (numberSplit.Any(string.IsNullOrWhiteSpace)) return null;
        if (numberSplit.Length < 3 && p.Class != ParagraphClass.Heading) return null;

        int titleBegin = numEnd;

        while (titleBegin < text.Length && !char.IsLetter(text[titleBegin]))
            titleBegin += 1;

        if (titleBegin == text.Length) return null;

        string title = text.Substring(titleBegin).Trim().TrimEnd('.');

        return new HeadingClassifierData
        {
            Number = number,
            Title = title
        };
    }
}