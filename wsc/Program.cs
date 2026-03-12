using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

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

root.SetAction(async res =>
{
    XmlTranslationsFile translations = XmlTranslationsFile.LoadEmbedded();

    if (res.GetValue(listDiagnosticsOpt))
    {
        LintManager manager = new LintManager();

        foreach (var diagnostic in manager.AllPossibleDiagnostics)
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

    LinterThreadPool pool = new(Environment.ProcessorCount);

        string suffix = autofix ? "FIXED" : "ANNOTATED";
        string target = Path.GetFileNameWithoutExtension(file.Name) + $"-{suffix}.docx";

        using (var linter = new DocumentLinter(file.FullName))
        {
            string? only = res.GetValue(onlyOpt);
            if (only != null)
            {
                var set = only.Split(",").ToHashSet();
                linter.LintIdFilter = lint => set.Contains(lint);
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
        LintTask task = new LintTask(x.Open(FileMode.Open, FileAccess.Read), lintIdFilter, false, null);
        
        pool.AddTask(task);

        using var linter = await task.Result;

        await Task.Yield();

        foreach (var message in linter.Diagnostics)
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

            if (!autofix)
            {
                linter.DocumentAnalysis.WriteComment(message, translations);
            }
        }

        bool changed = false;
        if (!autofix && linter.Diagnostics.Count > 0)
        {
            changed = true;
        }
        else if (linter.RunAutofixes())
        {
            changed = true;
        }

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

        if (changed)
        {
            linter.SaveTo(target);
            Console.WriteLine("Changes have been saved to " + target);
        }

        return linter.Diagnostics.Count();
    })
    .ToList();

    int diagnosticsTotal = (await Task.WhenAll(tasks)).Sum();
    
    return diagnosticsTotal > 0 ? 1 : 0;
});

return root.Parse(args).Invoke();