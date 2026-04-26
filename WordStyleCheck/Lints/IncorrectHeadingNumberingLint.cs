using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class IncorrectHeadingNumberingLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics => ["IncorrectHeadingNumbering"];

    public void Run(ILintContext ctx)
    {
        var elements = ctx.Document.AllParagraphs
            .Select(ctx.Document.GetTool)
            .Where(x => x.HeadingData is { IsConclusion: false })
            .ToList();

        List<int> levels = [];
        foreach (var element in elements)
        {
            int level = element.HeadingData!.Level;

            if (level > levels.Count)
            {
                for (int i = levels.Count; i < level; i++)
                {
                    levels.Add(1);
                }
            }
            else
            {
                while (level < levels.Count)
                    levels.RemoveAt(levels.Count - 1);

                levels[^1] += 1;
            }

            string number = string.Join(".", levels);

            if (number != element.HeadingData.Number)
            {
                // TODO: figure out whether this is a content or formatting error.
                ctx.AddMessage(new LintDiagnostic("IncorrectHeadingNumbering", DiagnosticType.ContentError, new ParagraphDiagnosticContext(element.Paragraph))
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
