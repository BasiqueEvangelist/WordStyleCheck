using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class NumberingPropertiesTool : INumbering
{
    private DocumentAnalysisContext _ctx;
    public NumberingInstance Numbering { get; }
    
    public AbstractNum AbstractNumbering { get; }

    internal NumberingPropertiesTool(DocumentAnalysisContext ctx, NumberingInstance numbering)
    {
        _ctx = ctx;
        Numbering = numbering;

        AbstractNumbering = ctx.Document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.ChildElements
            .OfType<AbstractNum>().First(x => x.AbstractNumberId!.Value == numbering.AbstractNumId!.Val!);
    }

    public List<Paragraph> Paragraphs { get; } = [];

    public string GetNumber(Paragraph p)
    {
        List<int> numbers = [];
        
        foreach (Paragraph sub in Paragraphs)
        {
            var sTool = _ctx.GetTool(sub);

            if (sTool.NumberingLevel + 1 > numbers.Count)
            {
                for (int i = numbers.Count; i <= sTool.NumberingLevel; i++)
                {
                    numbers.Add(AbstractNumbering.ChildElements.OfType<Level>().First(x => x.LevelIndex!.Value == i).StartNumberingValue?.Val ?? 0);
                }
            }
            else
            {
                while (sTool.NumberingLevel + 1 < numbers.Count)
                    numbers.RemoveAt(numbers.Count - 1);

                numbers[^1] += 1;
            }

            if (sub == p)
            {
                var level = AbstractNumbering.ChildElements.OfType<Level>().First(x => x.LevelIndex!.Value == numbers.Count - 1);

                string text = level.LevelText!.Val!.Value!;

                for (int i = 0; i < numbers.Count; i++)
                {
                    text = text.Replace($"%{i + 1}", numbers[i].ToString());
                }

                return text;
            }
        }

        throw new ArgumentException(nameof(p));
    }
}