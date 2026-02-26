using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck
{
    public class DocumentLinter : IDisposable
    {
        private string? _tempPath;
        private WordprocessingDocument? _document;
        private DocumentAnalysisContext _analysisCtx;
        private LintManager _manager;
        private LintContext _lintCtx;

        public DocumentLinter(string path, bool takeOwnership = false)
        {
            if (takeOwnership)
            {
                _tempPath = path;
            }
            else
            {
                _tempPath = Path.GetTempFileName();
                File.Copy(path, _tempPath, true);
            }

            _document = WordprocessingDocument.Open(_tempPath, true);
            _analysisCtx = new DocumentAnalysisContext(_document);
            _manager = new LintManager();
            _lintCtx = new LintContext(_analysisCtx, false);
        }
        
        public DocumentLinter(Stream stream)
        {
            _tempPath = Path.GetTempFileName();

            using (var wStream = File.OpenWrite(_tempPath))
                stream.CopyTo(wStream);

            _document = WordprocessingDocument.Open(_tempPath, true);
            _analysisCtx = new DocumentAnalysisContext(_document);
            _manager = new LintManager();
            _lintCtx = new LintContext(_analysisCtx, false);
        }

        public List<LintMessage> Diagnostics => _lintCtx.Messages;

        public Predicate<LintMessage> LintFilter
        {
            get => _lintCtx.LintFilter;
            set => _lintCtx.LintFilter = value;
        }

        public DocumentAnalysisContext DocumentAnalysis => _analysisCtx;

        public void RunLints()
        {
            _manager.Run(_lintCtx);
        }

        public void SaveTo(string path)
        {
            if (_tempPath == null) throw new InvalidOperationException("Document was already saved");

            _document.Save();
            _document.Dispose();
            _document = null;

            File.Move(_tempPath, path, true);
            _tempPath = null;
        }
        
        public string SaveTemp()
        {
            if (_document == null) return _tempPath!;

            _document.Save();
            _document.Dispose();
            _document = null;

            return _tempPath!;
        }

        public void Dispose()
        {
            _document?.Dispose();

            if (_tempPath != null) File.Delete(_tempPath);
        }

        public bool RunAutofixes()
        {
            return _lintCtx.RunAllAutoFixes();
        }

        public bool CanSave => _tempPath != null;
    }
}
