using Avalonia;
using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WordStyleCheck;

namespace WordStyleCheckGui.ViewModels;

public partial class DocumentViewModel : ViewModelBase
{
    private string _path;
    private string _fileName;
    private DocumentLinter _linter;
 
    public DocumentViewModel(string path)
    {
        _path = path;
        _fileName = Path.GetFileName(path);
        _linter = new DocumentLinter(path);

        _linter.RunLints();
    }

    public string FileName => _fileName;

    public string TotalDiagnosticsText => $"{_linter.Diagnostics.Count} style errors ({_linter.Diagnostics.Where(x => x.AutoFix != null).Count()} can be autofixed)";

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
            DefaultExtension = "docx"
        });

        if (file == null) return;

        // TODO: make this not an absolute path.
        var translations = DiagnosticTranslationsFile.LoadFromDocx("C:\\Users\\Nikolay\\source\\repos\\BasiqueEvangelist\\WordStyleCheck\\rules.docx");

        foreach (var message in _linter.Diagnostics)
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
        var translations = DiagnosticTranslationsFile.LoadFromDocx("C:\\Users\\Nikolay\\source\\repos\\BasiqueEvangelist\\WordStyleCheck\\rules.docx");

        _linter.RunAutofixes();

        _linter.SaveTo(file.TryGetLocalPath()!);
        CanSave = false;
    }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveAnnotatedCommand))]
    [NotifyCanExecuteChangedFor(nameof(SaveAutofixedCommand))]
    private bool _canSave = true;
}