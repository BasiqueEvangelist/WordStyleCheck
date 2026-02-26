using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace WordStyleCheckGui.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public ObservableCollection<DocumentViewModel> Documents { get; } = new();

    public bool ShowDndHint => Documents.Count == 0;

    public MainWindowViewModel()
    {
        Documents.CollectionChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(ShowDndHint));
        };
    }

    public void AddDocument(IStorageItem file)
    {
        Documents.Add(new DocumentViewModel(file.TryGetLocalPath()!));
    }

    [RelayCommand]
    private async Task OpenDialog()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions()
        {
            FileTypeFilter = [new("Microsoft Word files") {
                Patterns = ["*.docx"],
                // TODO: mime types
            }],
            AllowMultiple = true
        });

        if (files.Count == 0) return;

        foreach (var file in files)
        {
            Documents.Add(new DocumentViewModel(file.TryGetLocalPath()!));
        }
    }
}
