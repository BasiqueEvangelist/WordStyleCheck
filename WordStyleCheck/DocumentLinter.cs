using DocumentFormat.OpenXml.Packaging;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public class DocumentLinter : IDisposable
{
    private MemoryStream _stream;
    private WordprocessingDocument? _document;
    private DocumentAnalysisContext _analysisCtx;
    private LintManager _manager;
    private LintContext _lintCtx;

    public DocumentLinter(Stream stream)
    {
        if (stream is MemoryStream ms)
            _stream = ms;
        else
        {
            _stream = new MemoryStream();
            stream.CopyTo(_stream);
        }

        _document = WordprocessingDocument.Open(_stream, true);

        using (new LoudStopwatch("Loading document parts"))
        {
            _ = _document.MainDocumentPart!.Document;
            _ = _document.MainDocumentPart!.StyleDefinitionsPart!.Styles;
            _ = _document.MainDocumentPart!.NumberingDefinitionsPart?.Numbering;
            _ = _document.MainDocumentPart!.WordprocessingCommentsPart?.Comments;
        }
            
        using (new LoudStopwatch("StripOldComments.Run"))
            StripOldComments.Run(_document);
            
        _analysisCtx = new DocumentAnalysisContext(_document);
            
        _manager = new LintManager();
        _lintCtx = new LintContext(_analysisCtx, false);
    }
        
    public List<LintMessage> Diagnostics => _lintCtx.Messages;

    public Predicate<string> LintIdFilter
    {
        get => _lintCtx.LintIdFilter;
        set => _lintCtx.LintIdFilter = value;
    }

    public DocumentAnalysisContext DocumentAnalysis => _analysisCtx;

    public void RunLints()
    {
        _manager.Run(_lintCtx);
    }

    public void SaveTo(string path)
    {
        Save();
            
        _stream.Seek(0, SeekOrigin.Begin);

        using var fs = File.OpenWrite(path);
        _stream.CopyTo(fs);
        
        _stream.Seek(0, SeekOrigin.Begin);
    }
        
    public MemoryStream Save()
    {
        _stream.Seek(0, SeekOrigin.Begin);

        if (_document == null) return _stream!;

        if (!_document.AutoSave) _document.Save();
        _document.Dispose();
        _document = null;

        _stream.Seek(0, SeekOrigin.Begin);
            
        return _stream!;
    }

    public void Dispose()
    {
        _document?.Dispose();
    }

    public bool RunAutofixes()
    {
        return _lintCtx.RunAllAutoFixes();
    }
}