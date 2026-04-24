using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WordStyleCheck;
using WordStyleCheck.Profiles.Gost7_32;
using WordStyleCheckGui.Views;

namespace WordStyleCheckGui.ViewModels;

public partial class DocumentViewModel : ViewModelBase
{
    private static readonly LinterThreadPool Pool = new(Environment.ProcessorCount);
    private static readonly XmlTranslationsFile Translations = XmlTranslationsFile.LoadEmbedded();

    private readonly string _path;
    private readonly string _fileName;
    private DocumentLinter? _linter;

    private Exception? _exception;
 
    public DocumentViewModel(string path)
    {
        _path = path;
        _fileName = Path.GetFileName(path);
        
        async void RunThing()
        {
            var task = new LintTask(File.OpenRead(path), new Gost7_32Profile(), _ => true, false, null);
            Pool.AddTask(task);

            try
            {
                _linter = await task.Result;
            }
            catch (Exception e)
            {
                _exception = e;
            }

            Dispatcher.UIThread.Post(() =>
            {
                Done = true;
                CanSave = _exception == null;
            });
        }
        
        RunThing();
    }

    public string FileName => _fileName;

    public List<DiagnosticViewModel> Diagnostics => _linter == null ? [] : _linter.Diagnostics.Select(x => new DiagnosticViewModel(Translations, x)).ToList();

    public string TotalDiagnosticsText => _linter == null ? "" : $"{_linter.Diagnostics.Count} стилистических ошибок ({_linter.Diagnostics.Count(x => x.AutoFix != null)} автоисправляемых)";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(InProgress))]
    [NotifyPropertyChangedFor(nameof(IsError))]
    [NotifyPropertyChangedFor(nameof(Error))]
    [NotifyPropertyChangedFor(nameof(IsSuccess))]
    [NotifyPropertyChangedFor(nameof(TotalDiagnosticsText))]
    private bool _done;

    public bool InProgress => !Done;

    public bool IsSuccess => Done &&  _exception == null;
    
    public bool IsError => Done && _exception != null;

    public Exception? Error => _exception;

    [RelayCommand(CanExecute = nameof(CanSave))]
    private async Task SaveAnnotated()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        string suffix = "ANNOTATED";
        string target = Path.GetFileNameWithoutExtension(_path) + $"-{suffix}.docx";

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions()
        {
            SuggestedFileName = Path.GetFileName(target),
            SuggestedStartLocation = await storageProvider.TryGetFolderFromPathAsync(Path.GetDirectoryName(_path)!),
            FileTypeChoices = [MainWindowViewModel.DocxFileType],
            DefaultExtension = "docx"
        });

        if (file == null) return;

        foreach (var message in _linter!.Diagnostics)
        {
            _linter.DocumentAnalysis.WriteComment(message, Translations);
        }

        _linter.SaveTo(file.TryGetLocalPath()!);

        CanSave = false;
    }

    [RelayCommand(CanExecute = nameof(CanSave))]
    private async Task SaveAutofixed()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        string suffix = "FIXED";
        string target = Path.GetFileNameWithoutExtension(_path) + $"-{suffix}.docx";

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions()
        {
            SuggestedFileName = Path.GetFileName(target),
            SuggestedStartLocation = await storageProvider.TryGetFolderFromPathAsync(Path.GetDirectoryName(_path)!),
            DefaultExtension = "docx"
        });

        if (file == null) return;

        _linter!.RunAutofixes();

        _linter.SaveTo(file.TryGetLocalPath()!);
        CanSave = false;
    }

    [RelayCommand]
    private void OpenWindow()
    {
        var window = new DocumentReportWindow()
        {
            DataContext = this
        };

        window.Show();
    }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveAnnotatedCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveAutofixedCommand))]
    private bool _canSave = false;
}