using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectHeadingNumberingLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["IncorrectHeadingNumbering"];

    public void Run(LintContext ctx)
    {
        var elements = ctx.Document.AllParagraphs
            .Select(ctx.Document.GetTool)
            .Where(x => x.HeadingData is { IsConclusion: false })
            .ToList();

        List<int> levels = [];
        foreach (var element in elements)
        {
            int level = element.HeadingData!.Number.Count('.');

            if (level + 1 > levels.Count)
            {
                for (int i = levels.Count; i <= level; i++)
                {
                    levels.Add(1);
                }
            }
            else
            {
                while (level + 1 < levels.Count)
                    levels.RemoveAt(levels.Count - 1);

                levels[^1] += 1;
            }

            string number = string.Join(".", levels);

            if (number != element.HeadingData.Number)
            {
                ctx.AddMessage(new LintMessage("IncorrectHeadingNumbering", new ParagraphDiagnosticContext(element.Paragraph))
                {
                    Parameters = new()
                    {
                        ["Expected"] = number,
                        ["Actual"] = element.HeadingData.Number
                    }
                });
            }
        }
    }
}
