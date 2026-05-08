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
using WordStyleCheck.Profiles;
using WordStyleCheck.Profiles.Gost7_32;
using WordStyleCheckGui.Views;

namespace WordStyleCheckGui.ViewModels;

public partial class DocumentViewModel : ViewModelBase
{
    private static readonly LinterThreadPool Pool = new(Environment.ProcessorCount);

    private readonly MainWindowViewModel _mainWindow;
    private readonly string _path;
    private readonly string _fileName;
    private DocumentLinter? _linter;
    private readonly XmlTranslationsFile _translations;

    private Exception? _exception;
 
    public DocumentViewModel(MainWindowViewModel mainWindow, string path, string fileName, Stream stream)
    {
        _mainWindow = mainWindow;
        _path = path;
        _fileName = fileName;
        _translations = XmlTranslationsFile.LoadEmbedded(mainWindow.SelectedProfile.Id);
        
        async void RunThing()
        {
            var task = new LintTask(stream, mainWindow.SelectedProfile, _ => true, null, false);
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
                mainWindow.UpdateCompletedCount();
                CanSave = _exception == null;
            });
        }
        
        RunThing();
    }

    public string FileName => _fileName;

    public List<DiagnosticViewModel> Diagnostics => _linter == null ? [] : _linter.Diagnostics.Select(x => new DiagnosticViewModel(_translations, x)).ToList();

    public string TotalDiagnosticsText
    {
        get
        {
            if (_linter == null) return "";
            string s;

            if (_linter.FailedToOpen) s = "Невалидный файл";
            else if (_linter.SeriousError) s = "Найдены серьезные ошибки.";
            else s = $"{_linter.Diagnostics.Count} стилистических ошибок";

            return s;
        }
    }

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

    public string AnnotatedTarget => Path.GetFileNameWithoutExtension(_fileName) + "-ANNOTATED.docx";
    
    public MemoryStream SaveAnnotatedStream()
    {
        _linter!.ApplyDiagnostics(_translations);
        return _linter.Save();
    }
    
    [RelayCommand(CanExecute = nameof(CanSaveAnnotated))]
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

        _linter!.ApplyDiagnostics(_translations);

        _linter.SaveTo(file.TryGetLocalPath()!);

        CanSave = false;
    }

    [RelayCommand(CanExecute = nameof(CanSaveAutofixed))]
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

        _linter!.ClearDiagnostics();
        _linter.RunLints(true);
        _linter.ApplyDiagnostics(_translations);

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

    public bool CanSaveAnnotated => CanSave && !_linter!.FailedToOpen;
    public bool CanSaveAutofixed => CanSave && !_linter!.SeriousError;
}