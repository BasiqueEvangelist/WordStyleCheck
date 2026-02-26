using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

RootCommand root = new("Linter for .docx files");

Option<bool> generateRevisions = new("--revisions", "-r")
{
    Description = "Write changes as revisions to the document",
    DefaultValueFactory = _ => false
};

root.Options.Add(generateRevisions);

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

Argument<FileInfo> inputFile = new("input")
{
    Description = "File to lint for style issues",
};

root.Arguments.Add(inputFile);

root.SetAction(res =>
{
    if (res.GetValue(perfTimings))
        LoudStopwatch.Enabled = true;

    bool autofix = res.GetValue(autofixOpt);
    
    FileInfo input = res.GetValue(inputFile)!;
    
    // the OOXML toolkit we use doesn't seem to support saving to a different file, so we copy the document to a temporary
    // file, open it, and then copy it back to the -FIXED if it was changed in any way.

    String temp = Path.GetTempFileName();
    File.Copy(input.FullName, temp, true);

    List<LintMessage> messages;
    bool changed = false;
    int totalAutofixed = 0;

    var translations = DiagnosticTranslationsFile.LoadFromDocx("./rules.docx");

    using (var doc = WordprocessingDocument.Open(temp, true))
    {
        _ = doc.MainDocumentPart!.Document!;

        DocumentAnalysisContext analysisCtx = new(doc);
        
        LintContext ctx = new LintContext(analysisCtx, res.GetValue(generateRevisions));

        string? only = res.GetValue(onlyOpt);
        if (only != null)
        {
            var set = only.Split(",").ToHashSet();
            ctx.LintFilter = lint => set.Contains(lint.Id);
        }

        new LintManager().Run(ctx);
            
        foreach (var message in ctx.Messages)
        {
            Console.Write(Utils.ToPlainText(translations.Translate(message.Id, message.Parameters ?? new())));

            if (message.AutoFix != null && autofix)
            {
                Console.Write(" (autofixed)");
                totalAutofixed += 1;
            }

            Console.WriteLine(":");
        
            message.Context.WriteToConsole();

            if (!autofix)
            {
                analysisCtx.WriteComment(message, translations);
            }
        }

        if (!autofix && ctx.Messages.Count > 0)
        {
            changed = true;
        }
        else if (ctx.RunAllAutoFixes())
        {
            changed = true;
        }
        
        if (changed)
            doc.Save();

        messages = ctx.Messages;
    }

    if (messages.Count > 0)
    {
        Console.Write($"{messages.Count} style errors");

        if (totalAutofixed > 0)
        {
            Console.Write($" ({totalAutofixed} autofixed)");
        }

        Console.WriteLine(" :(");
    }
    else
    {
        Console.WriteLine("No errors detected :)");
    }

    if (changed)
    {
        string suffix = autofix ? "FIXED" : "ANNOTATED";
        string target = Path.GetFileNameWithoutExtension(input.Name) + $"-{suffix}.docx";
        File.Move(temp, target,true);
        Console.WriteLine("Changes have been saved to " + target);
    }
    else
    {
        File.Delete(temp);
    }

    return messages.Count > 0 ? 1 : 0;
});

return root.Parse(args).Invoke();