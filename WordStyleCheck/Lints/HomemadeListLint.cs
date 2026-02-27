using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class HomemadeListLint : ILint
{
    public void Run(LintContext ctx)
    {
        List<SniffedListData> possibleLists = [];
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            if (ctx.Document.GetTool(p).Class != ParagraphClass.BodyText) continue;
            
            string text = Utils.CollectParagraphText(p).Trim();

            if (text.StartsWith("—") || text.StartsWith("•"))
            {
                possibleLists.Add(new SniffedListData(true, [p]));
                continue;
            }

            if (text.Length > 0 && char.IsDigit(text[0]))
            {
                int i;
                for (i = 0; i < text.Length; i++)
                {
                    if (!char.IsDigit(text[i]))
                        break;
                }

                if (i != text.Length && (text[i] == '.' || text[i] == ')'))
                {
                    possibleLists.Add(new SniffedListData(false, [p]));
                    continue;
                }
            }
        }

        for (int i = 1; i < possibleLists.Count; i++)
        {
            if (possibleLists[i - 1].Paragraphs[^1].NextSibling() != possibleLists[i].Paragraphs[0]) continue;
            if (possibleLists[i - 1].Unordered != possibleLists[i].Unordered) continue;
            
            possibleLists[i - 1].Paragraphs.AddRange(possibleLists[i].Paragraphs);
            possibleLists.RemoveAt(i);
            i -= 1;
        }

        possibleLists.RemoveAll(x => x.Paragraphs.Count < 2);

        foreach (var list in possibleLists)
        {
            foreach (var p in list.Paragraphs)
            {
                ctx.Document.GetTool(p).OfNumbering = list;
            }
            
            ctx.AddMessage(new LintMessage("HandmadeList", new ParagraphDiagnosticContext(list.Paragraphs)));
        }
    }
    
    private record struct SniffedListData(bool Unordered, List<Paragraph> Paragraphs) : INumbering;
}