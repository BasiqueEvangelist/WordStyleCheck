using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace WordStyleCheckGui.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public static readonly FilePickerFileType DocxFileType = new("Файлы Microsoft Word") {
        Patterns = ["*.docx"],
        MimeTypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
        // TODO: mime types
    };

    public ObservableCollection<DocumentViewModel> Documents { get; } = new();

    public void AddDocument(IStorageItem file)
    {
        Documents.Add(new DocumentViewModel(file.TryGetLocalPath()!));
    }

    public async void OpenDialog()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions()
        {
            FileTypeFilter = [DocxFileType],
            AllowMultiple = true
        });

        if (files.Count == 0) return;

        foreach (var file in files)
        {
            Documents.Add(new DocumentViewModel(file.TryGetLocalPath()!));
        }
    }
}
