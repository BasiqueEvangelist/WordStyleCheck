using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace WscDebugTool.ViewModels;

public class XmlFileViewModel : ViewModelBase, IFileViewModel
{
    public string FileText { get; }
    
    public XmlFileViewModel(ZipArchiveEntry entry)
    {
        MemoryStream ms = new();

        using (var stream = entry.Open())
        {
            stream.CopyTo(ms);
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