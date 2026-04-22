using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class LintManager(IProfile profile)
{
    private readonly List<ILint> _lints = profile.Lints;
        
    public void Run(LintContext ctx)
    {
        foreach (var lint in _lints)
        {
            if (!lint.EmittedDiagnostics.Any(ctx.LintIdFilter.Invoke))
                continue;
            
            using (new LoudStopwatch(lint.GetType().Name))
            {
                try
                {
                    lint.Run(ctx);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Encountered exception while running {lint.GetType().Name}: {e}");
                }
            }
        }

        using (new LoudStopwatch("LintMerger.Run")) 
            LintMerger.Run(ctx.Messages);
    }

    public List<string> AllPossibleDiagnostics => _lints.SelectMany(x => x.EmittedDiagnostics).ToList();
}