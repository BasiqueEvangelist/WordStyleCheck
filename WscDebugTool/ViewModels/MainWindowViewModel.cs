using System;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
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
using DocumentFormat.OpenXml.Packaging;
using WordStyleCheck.Analysis;
using WordStyleCheck.Gost7_32;

namespace WscDebugTool.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShouldShowPickerButton))]
    [NotifyPropertyChangedFor(nameof(ArchiveReady))]
    [NotifyPropertyChangedFor(nameof(Files))]
    private ZipArchive? archive;

    private DocumentAnalysisContext? document;

    public bool ShouldShowPickerButton => Archive == null;
    public bool ArchiveReady => Archive != null;

    public List<FileEntryViewModel> Files => Archive == null ? [] : Archive.Entries.Select(x => new FileEntryViewModel(x, this)).ToList();

    [ObservableProperty] private IFileViewModel? _currentFile;
    
    [RelayCommand]
    private async Task OpenZipFile()
    {
        var storageProvider = ((IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!).MainWindow!.StorageProvider;

        var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions()
        {
            FileTypeFilter = [new("ZIP archives") {
                Patterns = ["*.zip", "*.docx", "*.odt"],
            }]
        });

        if (files == null || files.Count < 1) return;

        document = null;

        Archive = await ZipArchive.CreateAsync(await files[0].OpenReadAsync(), ZipArchiveMode.Read, false, null);

        if (files[0].Name.EndsWith(".docx"))
        {
            document = new DocumentAnalysisContext(WordprocessingDocument.Open(files[0].TryGetLocalPath()!, false), new Gost7_32Profile().Classifiers);
        } 
    }

    public void OpenFile(ZipArchiveEntry entry)
    {
        if (document != null)
        {
            var mainPartPath = document.Document.MainDocumentPart!.Uri.ToString().TrimStart('/');

            if (mainPartPath == entry.FullName)
            {
                CurrentFile = new DocumentFileViewModel(entry, document);
                return;
            }
        }

        if (entry.Name.EndsWith(".xml") || entry.Name.EndsWith(".rels"))
        {
            CurrentFile = new XmlFileViewModel(entry);
        } 
        else if (entry.Name.EndsWith(".png") || entry.Name.EndsWith(".jpg") || entry.Name.EndsWith(".jpeg"))
        {
            CurrentFile = new ImageFileViewModel(entry);
        }
    }
}
