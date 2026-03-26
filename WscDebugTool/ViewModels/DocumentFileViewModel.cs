using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WscDebugTool.ViewModels;

public class DocumentFileViewModel : XmlFileViewModel
{
    public DocumentAnalysisContext Document { get; }

    public List<ParagraphViewModel> Paragraphs =>
        Document.Document.MainDocumentPart!.Document!.Body!.ChildElements.OfType<Paragraph>().Select(x => new ParagraphViewModel(this, x)).ToList();
    
    private Dictionary<EasyXmlPath, FilePosition> elementMap = [];
    
    public DocumentFileViewModel(ZipArchiveEntry entry, DocumentAnalysisContext documentCtx) : base(entry)
    {
        Document = documentCtx;
        XmlReader reader = XmlReader.Create(new StringReader(FileText));
        IXmlLineInfo lineInfo = (IXmlLineInfo)reader;

        List<int> childIndices = [];
        
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (childIndices.Count > 0)
                {
                    var path = new EasyXmlPath([..childIndices]);
                    elementMap[path] = new FilePosition(lineInfo.LineNumber, lineInfo.LinePosition);
                }

                if (!reader.IsEmptyElement)
                {
                    childIndices.Add(0);
                }
                else
                {
                    childIndices[^1] += 1;
                }
            }
            
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                childIndices.RemoveAt(childIndices.Count - 1);
                
                if (childIndices.Count > 0) 
                    childIndices[^1] += 1;
            }
        }
    }

    private static EasyXmlPath GetPath(OpenXmlElement el)
    {
        List<int> indices = [];

        while (el.Parent != null)
        {
            indices.Add(el.Parent.ChildElements.ToList().IndexOf(el));
            el = el.Parent;
        }

        indices.Reverse();
        
        return new EasyXmlPath(indices);
    }

    public event Action<FilePosition> FocusParagraph; 
    
    public void Focus(Paragraph paragraph)
    {
        FocusParagraph?.Invoke(elementMap[GetPath(paragraph)]);
    }
}

public class EasyXmlPath(List<int> childIndices)
{
    public List<int> ChildIndices { get; } = childIndices;

    protected bool Equals(EasyXmlPath other)
    {
        return ChildIndices.SequenceEqual(other.ChildIndices);
    }

    public override bool Equals(object? obj)
    {
        if (obj is null) return false;
        if (ReferenceEquals(this, obj)) return true;
        if (obj.GetType() != GetType()) return false;
        return Equals((EasyXmlPath)obj);
    }

    public override int GetHashCode()
    {
        HashCode code = default;

        foreach (var value in childIndices)
        {
            code.Add(value.GetHashCode());
        }

        return code.ToHashCode();
    }
}

public record struct FilePosition(int Line, int Column);