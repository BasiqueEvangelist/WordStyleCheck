using System.CommandLine;
using WordStyleCheck;
using WordStyleCheck.Profiles;

RootCommand root = new("Linter for .docx files");

//Option<bool> generateRevisions = new("--revisions", "-r")
//{
//    Description = "Write changes as revisions to the document",
//    DefaultValueFactory = _ => false
//};

//root.Options.Add(generateRevisions);

Option<bool> perfTimings = new("--performance-timings", "-p")
{
    Description = "Dump performance timings to the console",
    DefaultValueFactory = _ => false
};

root.Options.Add(perfTimings);

Option<bool> autofixOpt = new("--fix", "-f")
{
    Description = "Automatically fix issues instead of writing comments",
    DefaultValueFactory = _ => false
};

root.Options.Add(autofixOpt);

Option<string?> onlyOpt = new("--only")
{
    Description = "Only invoke these lints",
};

root.Options.Add(onlyOpt);

Option<string?> ignoreOpt = new("--ignore")
{
    Description = "Ignore these lints",
};

root.Options.Add(ignoreOpt);

Argument<List<FileInfo>> inputFile = new("input")
{
    Description = "File to lint for style issues",
    Arity = ArgumentArity.OneOrMore
};

root.Arguments.Add(inputFile);

Option<bool> listDiagnosticsOpt = new("--list-diagnostics")
{
    Description = "List all possible diagnostics and their translations.",
    DefaultValueFactory = _ => false
};

root.Options.Add(listDiagnosticsOpt);

Option<bool> quietOpt = new("--quiet", "-q")
{
    Description = "Do not output all diagnostics.",
    DefaultValueFactory = _ => false
};

root.Options.Add(quietOpt);

Option<bool> debugReportOpt = new("--generate-debug-report", "-d")
{
    Description = "Generate debug reports for all files checked",
    DefaultValueFactory = _ => false
};

root.Options.Add(debugReportOpt);

Option<bool> onlyNewDiagnosticsOpt = new("--only-new-diagnostics")
{
    Description = "Use old debug report to only write new diagnostics",
    DefaultValueFactory = _ => false
};

root.Options.Add(onlyNewDiagnosticsOpt);

Option<string> profileOpt = new("--profile")
{
    Description = "Use this profile",
    DefaultValueFactory = _ => "gost-7.32"
};

root.Options.Add(profileOpt);

root.SetAction(async res =>
{
    string profileName = res.GetValue(profileOpt)!;
    XmlTranslationsFile translations = XmlTranslationsFile.LoadEmbedded(profileName);
    IProfile profile = ProfileStore.GetProfile(profileName)!;

    if (res.GetValue(listDiagnosticsOpt))
    {
        foreach (var diagnostic in profile.Lints.SelectMany(x => x.EmittedDiagnostics).Distinct())
        {
            Console.WriteLine($"{diagnostic}:");
            Console.WriteLine(Utils.ToPlainText(translations.Translate(diagnostic, new(), null)));
            Console.WriteLine();
        }
        
        return 0;
    }
    
    if (res.GetValue(perfTimings))
        LoudStopwatch.Enabled = true;

    bool autofix = res.GetValue(autofixOpt);
    
    List<FileInfo> input = res.GetValue(inputFile)!;

    int poolCount = Environment.ProcessorCount;
    LinterThreadPool pool = new(poolCount);

    Predicate<string> lintIdFilter = _ => true;
    
    string? only = res.GetValue(onlyOpt);
    if (only != null)
    {
        var set = only.Split(",").ToHashSet();
        lintIdFilter = lint => set.Contains(lint);
    }

    string? ignore = res.GetValue(ignoreOpt);
    if (ignore != null)
    {
        var set = ignore.Split(",").ToHashSet();
        lintIdFilter = lint => !set.Contains(lint);
    }
    
    string suffix = autofix ? "FIXED" : "ANNOTATED";
    
    List<Task<int>> tasks = input.Where(x =>
    {
        if (Path.GetFileNameWithoutExtension(x.Name).EndsWith("-ANNOTATED") ||
            Path.GetFileNameWithoutExtension(x.Name).EndsWith("-FIXED"))
        {
            Console.WriteLine($"Skipping {x.Name}");
            return false;
        }

        return true;
    })
    .Select(async x =>
    {
        string target = Path.GetFileNameWithoutExtension(x.Name) + $"-{suffix}.docx";
        LintTask task = new LintTask(x.Open(FileMode.Open, FileAccess.Read), profile, lintIdFilter, false, null);

        pool.AddTask(task);

        using var linter = await task.Result;

        await Task.Yield();

        string reportTarget = Path.GetFileNameWithoutExtension(x.Name) + "-REPORT.txt";

        HashSet<string> ids = [];

        if (res.GetValue(onlyNewDiagnosticsOpt) && File.Exists(reportTarget))
        {
            ids = File.ReadAllLines(reportTarget)
                .Select(x => x.Trim())
                .Where(x => x.StartsWith("--------") && x.EndsWith("--------"))
                .Select(x => x.Substring(8, x.Length - 16).Trim())
                .ToHashSet();
        }

        if (res.GetValue(debugReportOpt))
        {
            string reportTmp = Path.GetTempFileName();
            await using (var file = File.Open(reportTmp, FileMode.Create, FileAccess.Write, FileShare.Read))
            {
                await using StreamWriter sw = new(file);
                DebugReportGenerator report = new(sw);

                report.WriteHeader("WSC CLI tool");

                foreach (var diagnostic in linter.Diagnostics)
                {
                    report.WriteDiagnostic(diagnostic, translations);
                }

                sw.Flush();

                Console.WriteLine("Report has been saved to " + reportTarget);
            }
            
            File.Move(reportTmp, reportTarget, true);
        }

        var actualDiagnostics = linter.Diagnostics
            // TODO: add support for writing old comments that have vanished.
            .Where(y => ids.Count <= 0 || !ids.Contains(y.GetHash()))
            .ToList();
        
        foreach (var message in actualDiagnostics)
        {
            if (input.Count == 1  && !res.GetValue(quietOpt))
            {
                Console.Write(Utils.ToPlainText(translations.Translate(message.Id, message.Parameters ?? new(), null)));

                if (message.AutoFix != null && autofix)
                {
                    Console.Write(" (autofixed)");
                }

                Console.WriteLine(":");

                message.Context.WriteToConsole();
            }
        }

        bool changed = linter.ApplyDiagnostics(actualDiagnostics, translations, autofix);

        if (ids.Count > 0)
        {
            if (actualDiagnostics.Count > 0)
                Console.WriteLine($"{actualDiagnostics.Count} new style errors in {x.Name}");
        }
        else
        {
            if (linter.Diagnostics.Count > 0)
            {
                Console.Write($"{linter.Diagnostics.Count} style errors in {x.Name}");

                if (autofix)
                {
                    Console.Write($" ({linter.Diagnostics.Count(y => y.AutoFix != null)} autofixed)");
                }

                Console.WriteLine(" :(");
            }
            else
            {
                Console.WriteLine($"No errors detected in {x.Name} :)");
            }
        }

        if (changed)
        {
            linter.SaveTo(target);
            Console.WriteLine("Changes have been saved to " + target);
        }

        return linter.Diagnostics.Count;
    })
    .ToList();

    int diagnosticsTotal = (await Task.WhenAll(tasks)).Sum();
    
    return diagnosticsTotal > 0 ? 1 : 0;
});

return root.Parse(args).Invoke();