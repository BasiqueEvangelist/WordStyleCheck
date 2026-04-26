using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;
using WordStyleCheck.Profiles;

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
            
            string name = lint.GetType().Name;
            
            using (new LoudStopwatch(name))
            {
                try
                {
                    lint.Run(ctx);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Encountered exception while running {name}: {e}");
                    ctx.AddMessage(new LintDiagnostic(
                        "LintError",
                        DiagnosticType.Fatal,
                        new StartOfDocumentDiagnosticContext(),
                        new()
                        {
                            ["LintName"] = name,
                            ["Exception"] = e.ToString()
                        }
                    ));
                }
            }
        }

        using (new LoudStopwatch("LintMerger.Run")) 
            LintMerger.Run(ctx.Messages);
    }

    public List<string> AllPossibleDiagnostics => _lints.SelectMany(x => x.EmittedDiagnostics).ToList();
}