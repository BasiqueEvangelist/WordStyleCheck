using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;

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

    public ObservableCollection<DocumentViewModel> Documents { get; } = new();
    
    public async void AddDocument(string path)
    {
        if (path.EndsWith(".docx"))
        {
            Documents.Add(new DocumentViewModel(path, Path.GetFileName(path), File.OpenRead(path)));
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
                
                Documents.Add(new DocumentViewModel(entry.Name, entry.Name, ms));
            }
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
