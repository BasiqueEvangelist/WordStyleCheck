using System.Collections.ObjectModel;
using Avalonia.Platform.Storage;

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
        
    }
}
