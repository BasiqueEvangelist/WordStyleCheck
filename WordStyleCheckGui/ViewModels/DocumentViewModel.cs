using Avalonia;
using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WordStyleCheck;

namespace WordStyleCheckGui.ViewModels;

public partial class DocumentViewModel : ViewModelBase
{
    private readonly string _path;
    private readonly string _fileName;
    private DocumentLinter? _linter;
 
    public DocumentViewModel(string path)
    {
        _path = path;
        _fileName = Path.GetFileName(path);

        var thread = new Thread(() =>
        {
            _linter = new DocumentLinter(path);
            _linter.RunLints();

            Dispatcher.UIThread.Post(() =>
            {
                Done = true;
                CanSave = true;
            });
        })
        {
            Name = "Linter thread for " + _fileName
        };
        thread.Start();
    }

    public string FileName => _fileName;

    public string TotalDiagnosticsText => _linter == null ? "" : $"{_linter.Diagnostics.Count} стилистических ошибок ({_linter.Diagnostics.Count(x => x.AutoFix != null)} автоисправляемых)";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(InProgress))]
    [NotifyPropertyChangedFor(nameof(TotalDiagnosticsText))]
    private bool _done;

    public bool InProgress => !Done;

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

        var translations = DiagnosticTranslationsFile.LoadEmbedded();

        foreach (var message in _linter!.Diagnostics)
        {
            _linter.DocumentAnalysis.WriteComment(message, translations);
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

        // TODO: make this not an absolute path.
        var translations = DiagnosticTranslationsFile.LoadEmbedded();

        _linter!.RunAutofixes();

        _linter.SaveTo(file.TryGetLocalPath()!);
        CanSave = false;
    }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveAnnotatedCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveAutofixedCommand))]
    private bool _canSave = false;
}