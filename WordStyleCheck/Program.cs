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

Argument<FileInfo> inputFile = new("input")
{
    Description = "File to lint for style issues",
};

root.Arguments.Add(inputFile);

root.SetAction(res =>
{
    if (res.GetValue(perfTimings))
        LoudStopwatch.Enabled = true;
    
    FileInfo input = res.GetValue(inputFile)!;
    
    // the OOXML toolkit we use doesn't seem to support saving to a different file, so we copy the document to a temporary
    // file, open it, and then copy it back to the -FIXED if it was changed in any way.

    String temp = Path.GetTempFileName();
    File.Copy(input.FullName, temp, true);

    List<LintMessage> messages;
    bool changed = false;
    int totalAutofixed = 0;

    using (var doc = WordprocessingDocument.Open(temp, true))
    {
        _ = doc.MainDocumentPart!.Document!;

        DocumentAnalysisContext analysisCtx = new(doc);
        
        using (new LoudStopwatch("FieldStackTracker.Run"))
        {
            FieldStackTracker.Run(doc.MainDocumentPart!.Document!);
        }
        
        List<ILint> lints =
        [
            new NeedlessParagraphLint(),
            new ParagraphFirstLineIndentLint(),
            new ParagraphSpacingLint(),
            new BodyTextFontLint()
        ];
        LintContext ctx = new LintContext(analysisCtx, res.GetValue(generateRevisions));

        foreach (var lint in lints)
        {
            using (new LoudStopwatch(lint.GetType().Name))
            {
                lint.Run(ctx);
            }
        }

        using (new LoudStopwatch("RunLintMerger.Run")) 
            RunLintMerger.Run(ctx.Messages);
        
        using (new LoudStopwatch("ParagraphLintMerger.Run")) 
            ParagraphLintMerger.Run(ctx.Messages);


        foreach (var message in ctx.Messages)
        {
            Console.Write(message.Message);
        
            if (message.Values != null)
            {
                Console.Write($" (expected {message.Values?.Expected}, found {message.Values?.Actual})");
            }

            if (message.AutoFix != null)
            {
                Console.Write(" (autofixed)");
                totalAutofixed += 1;
            }

            Console.WriteLine(":");
        
            message.Context.WriteToConsole();
        }
        
        if (ctx.RunAllAutoFixes())
        {
            changed = true;
            doc.Save();
        }

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
        string target = Path.GetFileNameWithoutExtension(input.Name) + "-FIXED.docx";
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