using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;

namespace WordStyleCheck.Analysis;

public struct CaptionClassifierData
{
    public required CaptionType Type { get; init; }
    public required bool IsBelow { get; init; }
    public required bool IsContinuation { get; init; }
    public required string Number { get; init; }
    
    public required OpenXmlElement? TargetedElement { get; init; }
    public required StringSpan TypeSpan { get; init; }
    public required StringSpan NumberSpan { get; init; }
    
    public static CaptionClassifierData? Classify(ParagraphPropertiesTool p, bool secondPass)
    {
        CaptionType type;
        bool isBelow;
        OpenXmlElement? targeted;
        
        if (!secondPass)
        {
            if (p.Paragraph.Descendants<Picture>().Any())
            {
                type = CaptionType.Figure;
                targeted = null;

                isBelow = false;
                foreach (var desc in p.Paragraph.Descendants())
                {
                    if (desc is Picture)
                    {
                        isBelow = true;
                        break;
                    }

                    if (desc is Text t && !string.IsNullOrWhiteSpace(t.Text))
                    {
                        isBelow = false;
                        break;
                    }
                }

                if (!isBelow != secondPass)
                    return null;
            }
            else if (p.Paragraph.PreviousSibling() is Paragraph prev && prev.Descendants<Drawing>().Any())
            {
                type = CaptionType.Figure;
                targeted = p.Paragraph.PreviousSibling()!;
                isBelow = true;
            }
            else if (p.Paragraph.NextSibling() is Table)
            {
                type = CaptionType.Table;
                targeted = p.Paragraph.NextSibling()!;
                isBelow = false;
            }
            else
            {
                return null;
            }
        }
        else
        {
            if (p.Paragraph.NextSibling() is Paragraph next && next.Descendants<Drawing>().Any())
            {
                type = CaptionType.Figure;
                targeted = p.Paragraph.NextSibling()!;
                isBelow = false;
            }
            else if (p.Paragraph.PreviousSibling() is Table)
            {
                type = CaptionType.Table;
                targeted = p.Paragraph.PreviousSibling()!;
                isBelow = true;
            }
            else
            {
                return null;
            }
        }

        string text = p.Contents;
        
        int firstPartEnd;
        for (firstPartEnd = 0; firstPartEnd < text.Length; firstPartEnd++)
        {
            if (!(char.IsLetter(text[firstPartEnd]) || Utils.IsWhiteSpaceOrOtherJunk(text[firstPartEnd]) || text[firstPartEnd] == '.')) 
                break;
        }

        if (firstPartEnd >= text.Length) return null;

        string firstPart = Utils.StripJunk(text[..firstPartEnd].ToLowerInvariant());
        
        bool isContinuation = false;

        switch (type)
        {
            case CaptionType.Table when firstPart.Equals("продолжение таблицы", StringComparison.InvariantCultureIgnoreCase):
                isContinuation = true;        
                break;
            case CaptionType.Figure when !firstPart.Equals("рисунок", StringComparison.InvariantCultureIgnoreCase):
            case CaptionType.Table  when !firstPart.Equals("таблица", StringComparison.InvariantCultureIgnoreCase):
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
            Number = Utils.StripJunk(text.Substring(secondPartStart, secondPartEnd - secondPartStart)).TrimEnd('.'),
            
            TargetedElement = targeted,
            TypeSpan = new StringSpan(0, firstPartEnd),
            NumberSpan = new StringSpan(secondPartStart, secondPartEnd)
        };
    }

    public string GetCorrectText(string originalText)
    {
        originalText = originalText.Trim();
        
        int beginDesc = NumberSpan.EndExclusive;

        for (; beginDesc < originalText.Length; beginDesc++)
        {
            if (char.IsLetter(originalText[beginDesc]) || char.IsDigit(originalText[beginDesc]))
                break;
        }

        string desc = "";

        if (beginDesc < originalText.Length)
        {
            desc = originalText.Substring(beginDesc);
            desc = desc.Trim().TrimEnd('.');
        }

        if (IsContinuation)
            desc = "";

        string correct = (Type, IsContinuation) switch
        {
            (CaptionType.Figure, _) => "Рисунок ",
            (CaptionType.Table, false) => "Таблица ",
            (CaptionType.Table, true) => "Продолжение таблицы ",
            _ => throw new ArgumentOutOfRangeException()
        };

        correct += Number;

        if (desc != "")
        {
            correct += " – " + desc;
        }

        return correct;
    }
}

public enum CaptionType
{
    Figure,
    Table,
}

public record struct StringSpan(int BeginInclusive, int EndExclusive);