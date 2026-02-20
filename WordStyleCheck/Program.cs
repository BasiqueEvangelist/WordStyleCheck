using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck;
using WordStyleCheck.Lints;

RootCommand root = new("Linter for .docx files");

Option<bool> generateRevisions = new("--revisions", "-r")
{
    Description = "Write changes as revisions to the document",
    DefaultValueFactory = x => false
};

root.Options.Add(generateRevisions);

Argument<FileInfo> inputFile = new("input")
{
    Description = "File to lint for style issues",
};

root.Arguments.Add(inputFile);

root.SetAction(res =>
{
    FileInfo input = res.GetValue(inputFile)!;
    
    // the OOXML toolkit we use doesn't seem to support saving to a different file, so we copy the document to a temporary
    // file, open it, and then copy it back to the -FIXED if it was changed in any way.

    String temp = Path.GetTempFileName();
    File.Copy(input.FullName, temp, true);

    List<LintMessage> messages;
    bool changed = false;
    
    using (var doc = WordprocessingDocument.Open(temp, true))
    {
        List<ILint> lints =
        [
            new ParagraphFirstLineIndentLint(),
            new ParagraphSpacingLint(),
            new BodyTextFontLint()
        ];
        LintContext ctx = new LintContext(doc, true, res.GetValue(generateRevisions));

        foreach (var lint in lints)
        {
            lint.Run(ctx);
        }

        RunLintMerger.Run(ctx.Messages);
        
        if (ctx.DocumentChanged)
            doc.Save();

        changed = ctx.DocumentChanged;

        messages = ctx.Messages;
    }

    foreach (var message in messages)
    {
        Console.WriteLine(message.Message + (message.AutoFixed ? " (autofixed)" : "") + ":");
        message.Context.WriteToConsole();
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
});

return root.Parse(args).Invoke();