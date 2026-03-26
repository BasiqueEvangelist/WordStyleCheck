using System.IO;
using System.IO.Compression;
using Avalonia.Media.Imaging;

namespace WscDebugTool.ViewModels;

public class ImageFileViewModel : ViewModelBase, IFileViewModel
{
    public ImageFileViewModel(ZipArchiveEntry entry)
    {
        var ms = new MemoryStream();
        entry.Open().CopyTo(ms);
        ms.Seek(0, SeekOrigin.Begin);
        Image = new Bitmap(ms);
    }

    public Bitmap Image { get; }
}