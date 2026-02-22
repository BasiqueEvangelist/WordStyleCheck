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
        [StructuralElement.Bibliography] = ["СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ"],
        [StructuralElement.Appendix] = ["ПРИЛОЖЕНИЕ"]
    };

    public static string GetProperName(StructuralElement element)
    {
        return Names[element][0];
    }
    
    public static StructuralElement? Classify(Paragraph p)
    {
        string text = Utils.CollectParagraphText(p);

        if (text.Length > 100) return null;

        foreach (var entry in Names)
        {
            foreach (var nameOption in entry.Value)
            {
                if (Algorithms.LevenshteinNeighbors(text.ToUpperInvariant(), nameOption, 5))
                {
                    return entry.Key;
                }
            }
        }

        return null;
    }
}