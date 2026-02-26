using CommunityToolkit.Mvvm.ComponentModel;

namespace WordStyleCheckGui.ViewModels;

public partial class DocumentViewModel : ViewModelBase
{
    [ObservableProperty]
    private string _fileName = "";
}