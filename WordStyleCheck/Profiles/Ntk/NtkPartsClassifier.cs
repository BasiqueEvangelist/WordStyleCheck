using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Ntk;

public class NtkPartsClassifier : IClassifier
{
    [StringSyntax("Regex")]
    private const string NameRegexText = @"[А-Я][А-Я]+\s+[А-Я]\.\s*[А-Я]\.";
    
    private static readonly Regex NamesRegex = new Regex(@$"({NameRegexText},\s*)*{NameRegexText}");

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
        
        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (NamesRegex.IsMatch(tool.Contents)
                || tool.Contents.StartsWith("Научный руководитель")
                || tool.Contents.StartsWith("Научный консультант"))
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsAuthorData = true;
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
            else if (!tool.IsEmptyOrDrawing)
            {
                tool.GetFeature(NtkParagraphData.Key)!.IsProbablyJunk = true;
            }
            
            i++;
        }

        if (i >= paragraphs.Count) return;

        bool seenBibliography = false;
        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (!seenBibliography)
            {
                string content = Utils.StripJunk(tool.Contents).TrimEnd('.', ':');

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