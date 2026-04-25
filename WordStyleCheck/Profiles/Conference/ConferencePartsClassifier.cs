using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Conference;

public class ConferencePartsClassifier : IClassifier
{
    private static readonly Regex UdcRegex = new Regex("УДК ([0-9]+(?:\\.[0-9]+)*)");
    
    private static readonly Regex EmailRegex = new Regex("[\\w.!#$%&'*+/=?^`{|}~-]+@[a-z\\d](?:[a-z\\d-]{0,61}[a-z\\d])?(?:\\.[a-z\\d](?:[a-z\\d-]{0,61}[a-z\\d])?)*");
    
    public void Classify(DocumentAnalysisContext ctx)
    {
        var paragraphs = ctx.AllParagraphs.ToList();

        int i = 0;

        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (ParseUdc(tool) is { } udc)
            {
                tool.GetFeature(ConferenceParagraphData.Key)!.UniversalDecimalClassifier = udc;
                i++;
                break;
            }

            if (!tool.IsEmptyOrDrawing) break;
            
            i++;
        }

        if (i >= paragraphs.Count) return;
        
        // read authors until there are no more authors.
        
        while (i < paragraphs.Count)
        {
            var tool = ctx.GetTool(paragraphs[i]);

            if (tool.IsEmptyOrDrawing)
            {
                i += 1;
                continue;
            }
            
            string content = Utils.StripJunk(tool.Contents);

            if (content.Length < 1)
            {
                i += 1;
                continue;
            }

            if (!(char.IsLetter(content[0]) && char.IsUpper(content[0]) && !(content.Length >= 2 && char.IsUpper(content[1]))))
            {
                break;
            }

            if (tool.Justification == JustificationValues.Center) break;

            AuthorData author = new AuthorData();

            while (true)
            {
                author.Paragraphs.Add(tool);
                tool.GetFeature(ConferenceParagraphData.Key)!.AuthorData = author;

                i += 1;

                if (EmailRegex.Matches(tool.Contents).Count > 0)
                {
                    break;
                }

                if (i >= paragraphs.Count)
                    break;

                tool = ctx.GetTool(paragraphs[i]);

                if (tool.IsEmptyOrDrawing)
                {
                    break;
                }
            }
        }

        while (i < paragraphs.Count && ctx.GetTool(paragraphs[i]).IsEmptyOrDrawing) i += 1;
        
        if (i >= paragraphs.Count) return;

        ctx.GetTool(paragraphs[i]).GetFeature(ConferenceParagraphData.Key)!.IsTitle = true;

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
                        tool.GetFeature(ConferenceParagraphData.Key)!.IsBibliographyHeader = true;
                        seenBibliography = true;
                        break;
                    }
                }
            }
            else
            {
                if (!tool.Contents.StartsWith("©"))
                    tool.GetFeature(ConferenceParagraphData.Key)!.IsBibliographySource = true;
                else
                {
                    tool.GetFeature(ConferenceParagraphData.Key)!.IsCopyright = true;
                    break;
                }
            }

            i += 1;
        }
    }
    
    private static List<string>? ParseUdc(ParagraphPropertiesTool tool)
    {
        string text = Utils.StripJunk(tool.Contents);

        var matches = UdcRegex.Matches(text);
        if (matches.Count == 0) return null;

        return matches.Select(x => x.Groups[1].Captures[0].Value).ToList();
    }
}