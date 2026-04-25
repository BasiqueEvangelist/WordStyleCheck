using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Gost7_32;

public enum GostStructuralElement
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

public class GostStructuralElementClassifier : IClassifier
{
    private static readonly Dictionary<GostStructuralElement, List<string>> Names = new() {
        [GostStructuralElement.Authors] = ["СПИСОК ИСПОЛНИТЕЛЕЙ"],
        [GostStructuralElement.Abstract] = ["РЕФЕРАТ"],
        [GostStructuralElement.TableOfContents] = ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"],
        [GostStructuralElement.GlossaryTerms] = ["ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"], 
        [GostStructuralElement.GlossaryAbbreviations] = ["ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ"],
        [GostStructuralElement.Introduction] = ["ВВЕДЕНИЕ"], 
        [GostStructuralElement.Conclusion] = ["ЗАКЛЮЧЕНИЕ"],
        [GostStructuralElement.Bibliography] = Utils.BibliographyHeaderNames
    };

    public static string GetProperName(GostStructuralElement element)
    {
        return Names[element][0];
    }

    public void Classify(DocumentAnalysisContext ctx)
    {
        foreach (var p in ctx.AllParagraphs)
        {
            var tool = ctx.GetTool(p);
            
            if (tool is { IsTableOfContents: false, ContainingTableCell: null, NumberingId: null })
            {
                tool.GetFeature(GostParagraphData.Key)!.StructuralElementHeader = Classify(tool.Contents);
            }
        }
    }
    
    public static GostStructuralElement? Classify(string text)
    {
        if (text.Length > 100) return null;

        text = text.Trim().Trim('.', ':');
        
        if (ClassifyAppendixHeader(text))
            return GostStructuralElement.Appendix;

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
