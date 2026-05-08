using DocumentFormat.OpenXml.Packaging;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Profiles;

namespace WordStyleCheck;

public class DocumentLinter : IDisposable, ILintContext
{
    private MemoryStream _stream;
    private WordprocessingDocument? _document;
    private DocumentAnalysisContext? _analysisCtx;
    private IProfile _profile;

    private bool _autoFix;

    public DocumentLinter(Stream stream, IProfile profile)
    {
        if (stream is MemoryStream ms)
            _stream = ms;
        else
        {
            _stream = new MemoryStream();
            stream.CopyTo(_stream);
        }

        try
        {
            _document = WordprocessingDocument.Open(_stream, true);

            using (new LoudStopwatch("Loading document parts"))
            {
                _ = _document.MainDocumentPart!.Document!.Body;
                _ = _document.MainDocumentPart!.StyleDefinitionsPart!.Styles;
                _ = _document.MainDocumentPart!.NumberingDefinitionsPart?.Numbering;
                _ = _document.MainDocumentPart!.WordprocessingCommentsPart?.Comments;
            }
        }
        catch (Exception e)
        {
            ((ILintContext) this).AddMessage(new LintDiagnostic(
                "InvalidDocx",
                DiagnosticType.CouldNotOpen,
                new StartOfDocumentDiagnosticContext(),
                new()
                {
                    ["Exception"] = e.ToString()
                }
            ));

            _document = null;
        }

        if (_document != null)
        {
            using (new LoudStopwatch("StripOldComments.Run"))
                StripOldComments.Run(_document);

            _analysisCtx = new DocumentAnalysisContext(_document, profile.Classifiers);
        }

        _profile = profile;
    }
    
    public bool FailedToOpen => _analysisCtx == null;
    public bool SeriousError { get; private set; }

    public List<LintDiagnostic> Diagnostics { get; } = [];

    public Predicate<string> LintIdFilter { get; set; } = _ => true;

    public bool AutoFixed { get; private set; }

    bool ILintContext.GenerateRevisions => false;

    bool ILintContext.AutomaticallyFix => _autoFix;

    void ILintContext.AddMessage(LintDiagnostic diagnostic)
    {
        if (diagnostic.Type is DiagnosticType.CouldNotOpen or DiagnosticType.CouldNotParse)
        {
            SeriousError = true;
        }
        
        if (_autoFix) return;
        if (!LintIdFilter(diagnostic.Id)) return;
        
        Diagnostics.Add(diagnostic);
    }

    void ILintContext.MarkAutoFixed()
    {
        AutoFixed = true;
    }

    DocumentAnalysisContext ILintContext.Document => _analysisCtx!;

    public void RunLints(bool autoFix)
    {
        if (FailedToOpen) return;
        
        _autoFix = autoFix;
        
        foreach (var lint in _profile.Lints)
        {
            if (SeriousError) break;
            
            if (!lint.EmittedDiagnostics.Any(LintIdFilter.Invoke))
                continue;
            
            string name = lint.GetType().Name;
            
            using (new LoudStopwatch(name))
            {
                try
                {
                    lint.Run(this);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Encountered exception while running {name}: {e}");
                    ((ILintContext) this).AddMessage(new LintDiagnostic(
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

        if (!autoFix)
        {
            using (new LoudStopwatch("LintMerger.Run"))
                LintMerger.Run(Diagnostics);
        }

        _autoFix = false;
    }

    public void SaveTo(string path)
    {
        if (FailedToOpen) throw new InvalidOperationException();
        
        Save();
            
        _stream.Seek(0, SeekOrigin.Begin);

        string tmp = Path.GetTempFileName();
        using (var fs = File.OpenWrite(tmp))
        {
            _stream.CopyTo(fs);
        }
        File.Move(tmp, path, true);
        
        _stream.Seek(0, SeekOrigin.Begin);
    }
        
    public MemoryStream Save()
    {
        if (FailedToOpen) throw new InvalidOperationException();

        _stream.Seek(0, SeekOrigin.Begin);

        if (_document == null) return _stream;

        if (!_document.AutoSave) _document.Save();
        _document.Dispose();
        _document = null;

        _stream.Seek(0, SeekOrigin.Begin);
            
        return _stream;
    }

    public void Dispose()
    {
        _document?.Dispose();
    }

    public bool ApplyDiagnostics(List<LintDiagnostic> diagnostics, XmlTranslationsFile translations)
    {
        if (FailedToOpen) return false;
        
        bool changed = false;
        
        foreach (var message in diagnostics)
        {
            _analysisCtx!.WriteComment(message, translations);
            changed = true;
        }

        return changed;
    }
    
    public bool ApplyDiagnostics(XmlTranslationsFile translations)
    {
        return ApplyDiagnostics(Diagnostics, translations);
    }

    public void ClearDiagnostics()
    {
        if (FailedToOpen) return;
        
        Diagnostics.Clear();
        AutoFixed = false;
        SeriousError = false;
    }
}