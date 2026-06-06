using System.Text;
using WordStyleCheck.Analysis;

namespace WordStyleCheck;

public class RunAssociatedText
{
    public string Text { get; }

    public List<RunTextRun> Runs { get; }
    
    private RunAssociatedText(string text, List<RunTextRun> runs)
    {
        Text = text;
        Runs = runs;
    }

    public static RunAssociatedText FromParagraph(ParagraphPropertiesTool tool)
    {
        StringBuilder contents = new();
        List<RunTextRun> runs = new();
        
        foreach (var r in Utils.DirectRunChildren(tool.Paragraph))
        {
            int start = contents.Length;
            var runTool = tool.Context.GetTool(r);
            contents.Append(runTool.Contents);
            runs.Add(new RunTextRun(runTool, start));
        }

        return new RunAssociatedText(contents.ToString(), runs);
    }

    public RunPropertiesTool GetRunAt(int index)
    {
        foreach (var run in Runs)
        {
            if (run.StartIndex <= index && run.StartIndex + run.Run.Contents.Length > index)
            {
                return run.Run;
            }
        }

        throw new IndexOutOfRangeException();
    }

    public RunSpan GetSpan(int start, int length)
    {
        List<RunPropertiesTool> runs = [];
        int firstStart = -1;
        int lastEnd = -1;
        
        foreach (var run in Runs)
        {
            if (run.StartIndex >= start + length) break;
            if (run.StartIndex + run.Run.Contents.Length <= start) continue;

            runs.Add(run.Run);
            
            if (firstStart == -1)
            {
                firstStart = start - run.StartIndex;
            }

            if (start + length - run.StartIndex <= run.Run.Contents.Length)
            {
                lastEnd = start + length - run.StartIndex;
                break;
            }
        }

        return new RunSpan(runs, firstStart, lastEnd);
    }
}

public record struct RunTextRun(RunPropertiesTool Run, int StartIndex)
{
    public string GetPart(int start, int length)
    {
        return Run.Contents.Substring(
            Math.Max(0, start - StartIndex),
            Math.Min(start + length, StartIndex + Run.Contents.Length) - Math.Max(StartIndex, start)
        );
    }
}