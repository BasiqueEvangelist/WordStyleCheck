using System.IO.Hashing;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck;

public record class RunSpan(List<RunPropertiesTool> Runs, int FirstStart, int LastEnd)
{
    public void Hash(NonCryptographicHashAlgorithm hasher)
    {
        foreach (var r in Runs)
        {
            HashUtils.HashElement(r.Run, hasher);
        }
        hasher.Append(BitConverter.GetBytes(FirstStart));
        hasher.Append(BitConverter.GetBytes(LastEnd));
    }
    
    public override string ToString()
    {
        StringBuilder sb = new();

        for (int i = 0; i < Runs.Count; i++)
        {
            if (i == 0)
            {
                sb.Append(Runs[i].Contents[FirstStart..]);
            }
            else if (i == Runs.Count - 1)
            {
                sb.Append(Runs[i].Contents[..LastEnd]);
            }
            else
            {
                sb.Append(Runs[i].Contents);
            }
        }
        
        return sb.ToString();
    }
    
    public IEnumerable<Run> Isolate()
    {
        if (Runs.Count == 1)
        {
            int len = Runs[0].Contents.Length;
            var center = Runs[0].Run;
            
            if (FirstStart != 0)
            {
                var split = Utils.SplitRunAt(center, FirstStart);

                center = split.second;
            }
            
            if (LastEnd != len)
            {
                center = Utils.SplitRunAt(center, LastEnd - FirstStart).first;
            }

            yield return center;

            yield break;
        }

        for (var i = 0; i < Runs.Count; i++)
        {
            if (i == 0 && FirstStart != 0)
            {
                var split = Utils.SplitRunAt(Runs[i].Run, FirstStart);
                yield return split.second;
            }
            else if (i == Runs.Count - 1 && LastEnd != Runs[i].Contents.Length)
            {
                var split = Utils.SplitRunAt(Runs[i].Run, LastEnd);
                yield return split.first;
            }
            else
            {
                yield return Runs[i].Run;
            }
        }
    }

    public void Replace(string contents)
    {
        var properties = (RunProperties?) Runs.MaxBy(x => x.Contents.Length)!.Run.RunProperties?.CloneNode(true);
        var runs = Isolate().ToList();

        var run = new Run();
        run.RunProperties = properties;
        run.AppendChild(new Text(contents)
        {
            Space = SpaceProcessingModeValues.Preserve
        });

        runs[0].InsertBeforeSelf(run);

        foreach (var old in runs)
        {
            old.Remove();
        }
    }

    public bool Matches(Predicate<RunPropertiesTool> predicate)
    {
        for (var i = 0; i < Runs.Count; i++)
        {
            string part;
            if (i == 0)
            {
                part = Runs[i].Contents[FirstStart..];
            }
            else if (i == Runs.Count - 1)
            {
                part = Runs[i].Contents[..LastEnd];
            }
            else
            {
                part = Runs[i].Contents;
            }
            
            if (string.IsNullOrWhiteSpace(part)) continue;
            
            if (!predicate(Runs[i])) return false;
        }

        return true;
    }
}