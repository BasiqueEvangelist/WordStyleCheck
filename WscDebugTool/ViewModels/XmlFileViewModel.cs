using CommunityToolkit.Mvvm.Input;
using System.IO.Compression;
using System.Threading.Tasks;

namespace WscDebugTool.ViewModels;

public partial class XmlFileViewModel(ZipArchiveEntry zipEntry, MainWindowViewModel mainWindow) : ViewModelBase
{
    public string Name { get; } = zipEntry.FullName;

    [RelayCommand]
    private Task OpenFile()
    {
        return mainWindow.OpenFile(zipEntry);
    }
}