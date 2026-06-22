using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Ntk;

public class NtkPartsClassifier : IClassifier
{
    [StringSyntax("Regex")]
    public const string NameRegexText = @"[А-Я][а-я]+\s+[А-Я]\.\s*[А-Я]\.";

    [StringSyntax("Regex")]
    private const string LenientNameRegexText = @"[А-Я][а-я]+\s+[А-Я](\.|([а-я]+))\s*[А-Я](\.|([а-я]+))";
    
    private static readonly Regex NamesRegex = new Regex(@$"({LenientNameRegexText},\s*)*{LenientNameRegexText}");

    private static readonly SearchValues<char> RussianLetters = SearchValues.Create("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя");
    private static readonly SearchValues<char> LatinLetters = SearchValues.Create("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz");

    public void Classify(DocumentAnalysisContext ctx)
    {
        var paragraphs = ctx.AllParagraphs.ToList();

        int i = 0;

        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (ParseUdc(tool))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsUniversalDecimalClassifier = true;
                i++;
                break;
            }

            if (!tool.IsEmptyOrDrawing) break;
            
            i++;
        }

        while (i < paragraphs.Count && ctx.GetTool(paragraphs[i]).IsEmptyOrDrawing) i += 1;
        
        if (i >= paragraphs.Count) return;

        ctx.GetTool(paragraphs[i]).GetFeature(NtkParagraphData.Key)!.IsTitle = true;
        i++;

        bool attachedAuthor = false;
        
        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (tool.Contents.Contains("руководитель", StringComparison.InvariantCultureIgnoreCase))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsSupervisorData = true;
            }
            else if (tool.Contents.StartsWith("консультант", StringComparison.InvariantCultureIgnoreCase))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsConsultantData = true;
            }
            else if (NamesRegex.IsMatch(tool.Contents) && !attachedAuthor)
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsAuthorData = true;
                attachedAuthor = true;
            }
            else if (tool.Contents.StartsWith("Аннотация"))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsAbstract = true;
            }
            else if (tool.Contents.StartsWith("Ключевые слова"))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsKeywords = true;
                i++;
                break;
            }
            else if (tool.Contents.StartsWith("РТУ МИРЭА,"))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsSourceInstitute = true;
            }
            else if (!tool.IsEmptyOrDrawing)
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsProbablyJunk = true;
            }
            
            i++;
        }

        if (i >= paragraphs.Count) return;

        bool seenBibliography = false;
        bool seenNormalText = false;
        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (!seenBibliography)
            {
                string content = Utils.StripJunk(tool.Contents).TrimEnd('.', ':');
                
                if (!seenNormalText)
                {
                    if (MemoryExtensions.ContainsAny(content, RussianLetters))
                    {
                        seenNormalText = true;
                    }
                    else if (MemoryExtensions.ContainsAny(content, LatinLetters))
                    {
                        tool.GetFeature(NtkParagraphData.Key)!.IsProbablyJunk = true;
                    }
                }

                if (!string.IsNullOrWhiteSpace(content))
                {
                    if (tool.Runs.All(x => x.Bold || string.IsNullOrWhiteSpace(x.Contents)) || tool.Runs.All(x => x.Italic || string.IsNullOrWhiteSpace(x.Contents)))
                        tool.GetFeature(NtkParagraphData.Key)!.IsProbablyHeading = true;
                }

                foreach (var option in Utils.BibliographyHeaderNames)
                {
                    if (content.Equals(option, StringComparison.InvariantCultureIgnoreCase))
                    {
                        tool.GetFeature(NtkParagraphData.Key)!.IsBibliographyHeader = true;
                        seenBibliography = true;
                        break;
                    }
                }
            }
            else
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsBibliographySource = true;
            }

            i += 1;
        }
    }
    
    private static bool ParseUdc(ParagraphPropertiesTool tool)
    {
        string text = Utils.StripJunk(tool.Contents);

        return text.Contains("УДК");
    }
}