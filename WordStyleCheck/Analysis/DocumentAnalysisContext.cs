using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class DocumentAnalysisContext
{
    public WordprocessingDocument Document { get; }

    private readonly Dictionary<Paragraph, ParagraphPropertiesTool> _paragraphTools = new();
    private readonly Dictionary<Run, RunPropertiesTool> _runTools = new();

    private readonly Dictionary<string, Style> _styles = new();

    private readonly List<SectionPropertiesTool> _sections = [];

    public Style? DefaultParagraphStyle { get; }

    public DocumentAnalysisContext(WordprocessingDocument document)
    {
        Document = document;

        if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
        {
            _styles = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements.OfType<Style>()
                .ToDictionary(x => x.StyleId!.Value!, x => x);

            DefaultParagraphStyle = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements.OfType<Style>()
                .SingleOrDefault(x => x.Type?.Value == StyleValues.Paragraph && (x.Default?.Value ?? false));
        }

        FieldStackTracker.Run(document.MainDocumentPart!.Document!);
        
        AllParagraphs = Document.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().ToList();

        StructuralElement? currentElement = null;
        List<Paragraph> currentSection = [];
        
        foreach (var p in AllParagraphs)
        {
            var tool = GetTool(p);

            if (tool.StructuralElementHeader != null)
            {
                currentElement = tool.StructuralElementHeader;
            }

            tool.OfStructuralElement = currentElement;
            
            currentSection.Add(p);

            if (p.ParagraphProperties?.SectionProperties is { } sectPr)
            {
                SectionPropertiesTool section = new(this, sectPr)
                {
                    Paragraphs = currentSection
                };
                _sections.Add(section);
                
                currentSection = [];
            }
        }

        if (Document.MainDocumentPart!.Document!.Body!.LastChild is SectionProperties lastSectPr)
        {
            SectionPropertiesTool section = new(this, lastSectPr)
            {
                Paragraphs = currentSection
            };
            _sections.Add(section);
        }
        else
        {
            throw new NotImplementedException("No last section in document?");
        }
    }

    public ParagraphPropertiesTool GetTool(Paragraph p)
    {
        if (_paragraphTools.TryGetValue(p, out var tool)) return tool;
        
        tool = new ParagraphPropertiesTool(this, p);
        _paragraphTools[p] = tool;

        return tool;
    }
    
    public RunPropertiesTool GetTool(Run r)
    {
        if (_runTools.TryGetValue(r, out var tool)) return tool;
        
        tool = new RunPropertiesTool(this, r);
        _runTools[r] = tool;

        return tool;
    }

    public Style? GetStyle(string styleId)
    {
        return _styles.GetValueOrDefault(styleId);
    }

    public T? FollowStyleChain<T>(string? styleId, Func<Style, T?> getter)
    {
        while (styleId != null)
        {
            var style = GetStyle(styleId);

            if (style == null) return default;

            var result = getter(style);
            if (result != null) return result;

            styleId = style.BasedOn?.Val?.Value;
        }

        return default;
    }

    public bool SniffStyleName(string? styleId, string name)
    {
        while (styleId != null)
        {
            var style = GetStyle(styleId);

            if (style == null) break;

            if ((style.StyleId?.InnerText?.Contains(name, StringComparison.InvariantCultureIgnoreCase) ?? false)
                || (style.StyleName?.InnerText.Contains(name, StringComparison.InvariantCultureIgnoreCase) ?? false))
            {
                return true;
            }

            styleId = style.BasedOn?.Val?.Value;
        }
        
        return false;
    }

    private int commentId = 0;
    
    public void WriteComment(LintMessage msg, DiagnosticTranslationsFile translations)
    {
        string id = (commentId++).ToString();
        
        msg.Context.WriteCommentReference(id);

        var commentsPart = Document.MainDocumentPart!.WordprocessingCommentsPart ??
            Document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();

        commentsPart.Comments ??= new Comments();

        // TODO: extract from .docx of lint messages
        Comment c = new Comment()
        {
            Id = id,
            Author = "WordStyleCheck",
            Initials = "WSC",
            Date = DateTime.Now
        };

        if (translations.Translations.TryGetValue(msg.Id, out var translation))
        {
            if (msg.Parameters != null)
            {
                foreach (var entry in msg.Parameters.OrderByDescending(x => x.Key))
                {
                    translation = translation.Replace("$" + entry.Key, entry.Value);
                }
            }

            c.InnerXml = translation;
        } else
        {
            c.Append(new Paragraph(new Run(new Text(msg.Id))));
        }

        commentsPart.Comments.AppendChild(c);
        commentsPart.Comments.Save();
    }

    public IReadOnlyList<SectionPropertiesTool> AllSections => _sections;
    public IEnumerable<Paragraph> AllParagraphs { get; }
}