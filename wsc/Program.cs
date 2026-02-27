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

root.SetAction(res =>
{
    if (res.GetValue(perfTimings))
        LoudStopwatch.Enabled = true;

    bool autofix = res.GetValue(autofixOpt);
    
    List<FileInfo> input = res.GetValue(inputFile)!;

    int diagnosticsTotal = 0;

    foreach (var file in input)
    {
        if (Path.GetFileNameWithoutExtension(file.Name).EndsWith("-ANNOTATED") ||
            Path.GetFileNameWithoutExtension(file.Name).EndsWith("-FIXED"))
        {
            Console.WriteLine($"Skipping {file.Name}");
            continue;
        }
        
        Console.WriteLine($"Processing {file.Name}");
        // the OOXML toolkit we use doesn't seem to support saving to a different file, so we copy the document to a temporary
        // file, open it, and then copy it back to the -FIXED if it was changed in any way.

        String temp = Path.GetTempFileName();
        File.Copy(file.FullName, temp, true);

        var translations = DiagnosticTranslationsFile.LoadEmbedded();

        string suffix = autofix ? "FIXED" : "ANNOTATED";
        string target = Path.GetFileNameWithoutExtension(file.Name) + $"-{suffix}.docx";

        using (var linter = new DocumentLinter(file.FullName))
        {
            string? only = res.GetValue(onlyOpt);
            if (only != null)
            {
                var set = only.Split(",").ToHashSet();
                linter.LintFilter = lint => set.Contains(lint.Id);
            }

            string? ignore = res.GetValue(ignoreOpt);
            if (ignore != null)
            {
                var set = ignore.Split(",").ToHashSet();
                linter.LintFilter = lint => !set.Contains(lint.Id);
            }

            linter.RunLints();

            foreach (var message in linter.Diagnostics)
            {
                if (input.Count == 1)
                {
                    Console.Write(Utils.ToPlainText(translations.Translate(message.Id, message.Parameters ?? new())));

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
                Console.Write($"{linter.Diagnostics.Count} style errors");

                if (autofix)
                {
                    Console.Write($" ({linter.Diagnostics.Count(x => x.AutoFix != null)} autofixed)");
                }

                Console.WriteLine(" :(");
            }
            else
            {
                Console.WriteLine("No errors detected :)");
            }

            if (changed)
            {
                linter.SaveTo(target);
                Console.WriteLine("Changes have been saved to " + target);
            }

            diagnosticsTotal += linter.Diagnostics.Count;
        }
    }
    
    return diagnosticsTotal > 0 ? 1 : 0;
});

return root.Parse(args).Invoke();