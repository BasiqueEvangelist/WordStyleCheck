using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public enum StructuralElement
{
    Authors,
    Abstract,
    TableOfContents,
    GlossaryTerms,
    GlossaryAbbreviations,
    Introduction,
    Conclusion,
    Bibliography,
    Appendix
}

public static class StructuralElementHeaderClassifier
{
    private static readonly Dictionary<StructuralElement, List<string>> Names = new() {
        [StructuralElement.Authors] = ["СПИСОК ИСПОЛНИТЕЛЕЙ"],
        [StructuralElement.Abstract] = ["РЕФЕРАТ"],
        [StructuralElement.TableOfContents] = ["СОДЕРЖАНИЕ"],
        [StructuralElement.GlossaryTerms] = ["ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"], 
        [StructuralElement.GlossaryAbbreviations] = ["ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ"],
        [StructuralElement.Introduction] = ["ВВЕДЕНИЕ"], 
        [StructuralElement.Conclusion] = ["ЗАКЛЮЧЕНИЕ"],
        [StructuralElement.Bibliography] = ["СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ", "СПИСОК ИСТОЧНИКОВ И ЛИТЕРАТУРЫ"]
    };

    public static string GetProperName(StructuralElement element)
    {
        return Names[element][0];
    }
    
    public static StructuralElement? Classify(Paragraph p)
    {
        string text = Utils.CollectParagraphText(p);

        if (text.Length > 100) return null;

        if (ClassifyAppendixHeader(text))
            return StructuralElement.Appendix;

        foreach (var entry in Names)
        {
            foreach (var nameOption in entry.Value)
            {
                if (text.Equals(nameOption, StringComparison.InvariantCultureIgnoreCase))
                {
                    return entry.Key;
                }
            }
        }

        return null;
    }

    public static bool ClassifyAppendixHeader(string text)
    {
        int firstPartEnd;
        for (firstPartEnd = 0; firstPartEnd < text.Length; firstPartEnd++)
        {
            if (!char.IsLetter(text[firstPartEnd])) 
                break;
        }
        
        if (firstPartEnd >= text.Length) return false;

        string firstPart = text[..firstPartEnd].ToLowerInvariant();
        
        return firstPart.Equals("приложение", StringComparison.InvariantCultureIgnoreCase);
    }
}