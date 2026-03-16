using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using AvaloniaEdit;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace WscDebugTool.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShouldShowPickerButton))]
    [NotifyPropertyChangedFor(nameof(ArchiveReady))]
    [NotifyPropertyChangedFor(nameof(Files))]
    private ZipArchive? archive;

    public bool ShouldShowPickerButton => Archive == null;
    public bool ArchiveReady => Archive != null;

    public List<XmlFileViewModel> Files => Archive == null ? [] : Archive.Entries.Select(x => new XmlFileViewModel(x, this)).ToList();

    [ObservableProperty]
    private string fileText = "Select a file";
    
    [RelayCommand]
    private async Task OpenZipFile()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions()
        {
            FileTypeFilter = [new("Файлы Microsoft Word") {
                Patterns = ["*.docx"],
                MimeTypes = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"]
            }]
        });

        if (files == null || files.Count < 1) return;

        Archive = await ZipArchive.CreateAsync(await files[0].OpenReadAsync(), ZipArchiveMode.Read, false, null);
    }

    public async Task OpenFile(ZipArchiveEntry entry)
    {
        MemoryStream ms = new();

        await using (var stream = await entry.OpenAsync())
        {
            await stream.CopyToAsync(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        string text;
        using (StreamReader reader = new StreamReader(ms,
                                              detectEncodingFromByteOrderMarks: true))
        {
            text = reader.ReadToEnd();
        }

        XDocument doc = XDocument.Parse(text);
        StringBuilder formatted = new();
        XmlWriter w = XmlWriter.Create(formatted, new XmlWriterSettings
        {
            Indent = true
        });
        doc.WriteTo(w);
        w.Flush();
        FileText = formatted.ToString();


    }
}
