using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace WordStyleCheckGui.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public static readonly FilePickerFileType DocxFileType = new("Файлы Microsoft Word") {
        Patterns = ["*.docx"],
        MimeTypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
    };
    
    public static readonly FilePickerFileType ZipFileType = new("ZIP архивы") {
        Patterns = ["*.zip"],
        MimeTypes = ["application/zip"]
    };

    [ObservableProperty, NotifyCanExecuteChangedFor(nameof(SaveAllDialogCommand))] private int _completedCount;
    [ObservableProperty, NotifyCanExecuteChangedFor(nameof(ClearCommand))] private bool _hasAny;

    public bool CanSaveAll => CompletedCount == Documents.Count && Documents.Count > 0;

    public ObservableCollection<DocumentViewModel> Documents { get; } = new();

    public void UpdateCompletedCount()
    {
        CompletedCount = Documents.AsEnumerable().Count(x => x.Done);
    }
    
    public async void AddDocument(string path)
    {
        await Task.Yield();

        if (path.EndsWith(".docx"))
        {
            Documents.Add(new DocumentViewModel(this, path, Path.GetFileName(path), File.OpenRead(path)));
            HasAny = true;
        } 
        else if (path.EndsWith(".zip"))
        {
            await using var archive = await ZipArchive.CreateAsync(File.OpenRead(path), ZipArchiveMode.Read, false, null);
            foreach (var entry in archive.Entries)
            {
                if (!entry.Name.EndsWith(".docx")) continue;

                MemoryStream ms = new();
                await using (var fs = await entry.OpenAsync())
                    await fs.CopyToAsync(ms);

                ms.Seek(0, SeekOrigin.Begin);
                
                Documents.Add(new DocumentViewModel(this, entry.Name, entry.Name, ms));
                HasAny = true;
            }
        }
    }

    [RelayCommand(CanExecute = nameof(HasAny))]
    private void Clear()
    {
        Documents.Clear();
        HasAny = false;
        UpdateCompletedCount();
    }

    [RelayCommand(CanExecute = nameof(CanSaveAll))]
    private async Task SaveAllDialog()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions()
        {
            FileTypeChoices = [ZipFileType]
        });

        if (file == null) return;

        await using var arch = await ZipArchive.CreateAsync(await file.OpenWriteAsync(), ZipArchiveMode.Create, false, null);

        foreach (var doc in Documents.Where(x => x.IsSuccess))
        {
            if (doc.Diagnostics.Count == 0) continue;

            var entry = arch.CreateEntry(doc.AnnotatedTarget);

            await using var stream = await entry.OpenAsync();
            await doc.SaveAnnotatedStream().CopyToAsync(stream);
        }
    } 

    public async void OpenDialog()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions()
        {
            FileTypeFilter = [DocxFileType, ZipFileType],
            AllowMultiple = true
        });

        if (files.Count == 0) return;

        foreach (var file in files)
        {
            AddDocument(file.TryGetLocalPath()!);
        }
    }
}
