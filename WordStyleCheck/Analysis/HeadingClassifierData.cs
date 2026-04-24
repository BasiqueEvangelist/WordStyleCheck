namespace WordStyleCheck.Analysis;

public class HeadingClassifierData
{
    public required string Number { get; init; }
    public required int Level { get; init; }
    public required string Title { get; init; }
    public required bool IsConclusion { get; init; }
    
    public static HeadingClassifierData? Classify(ParagraphPropertiesTool p)
    {
        if (p is { OfNumbering: not null, ProbablyHeading: false }) return null;
        if (p.IsEmptyOrDrawing) return null;
        if (p.ProbablyCodeListing || p.IsTableOfContents || p.ContainingTableCell != null || p.ContainingTextBox != null || p.EquationData != null) return null;
        
        string text = p.Contents.Trim();

        if (text.Length < 1) return null;

        int numEnd;
        string number;
        bool isConclusion = false;
        if (text.StartsWith("Выводы", StringComparison.InvariantCultureIgnoreCase) && p.ProbablyHeading)
        {
            isConclusion = true;
            number = "";
            numEnd = 0;
        }
        else if (p.OfNumbering is NumberingPropertiesTool numbering)
        {
            numEnd = 0;
            number = numbering.GetNumber(p.Paragraph).TrimEnd('.');
        }
        else
        {
            numEnd = 0;

            while (numEnd < text.Length && (char.IsDigit(text[numEnd]) || text[numEnd] == '.'))
                numEnd += 1;

            if (numEnd == text.Length) return null;

            number = text[..numEnd].TrimEnd('.');
        }

        int level;
        if (!isConclusion)
        {
            string[] numberSplit = number.Split('.');

            if (numberSplit.Any(string.IsNullOrWhiteSpace)) return null;
            if (numberSplit.Length < 3 && !p.ProbablyHeading) return null;

            level = numberSplit.Length;
        }
        else
        {
            level = -1; // TODO: calculate properly.
        }

        int titleBegin = numEnd;

        while (titleBegin < text.Length && !(char.IsLetter(text[titleBegin]) || char.IsDigit(text[titleBegin])))
            titleBegin += 1;

        if (titleBegin == text.Length) return null;

        string title = text.Substring(titleBegin).Trim().TrimEnd('.');

        return new HeadingClassifierData
        {
            Number = number,
            Level = level,
            Title = title,
            IsConclusion = isConclusion
        };
    }
}