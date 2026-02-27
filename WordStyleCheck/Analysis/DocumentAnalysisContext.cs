using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class DocumentAnalysisContext
{
    public WordprocessingDocument Document { get; }

    private readonly Dictionary<Paragraph, ParagraphPropertiesTool> _paragraphTools = new();
    private readonly Dictionary<Run, RunPropertiesTool> _runTools = new();

    private readonly Dictionary<(StyleValues, string), Style> _styles = new();

    private readonly List<SectionPropertiesTool> _sections = [];
    private readonly List<NumberingPropertiesTool> _numberings = [];

    public Style? DefaultParagraphStyle { get; }

    public DocumentAnalysisContext(WordprocessingDocument document)
    {
        Document = document;

        if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
        {
            _styles = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements.OfType<Style>()
                .ToDictionary(x => (x.Type!.Value!, x.StyleId!.Value!), x => x);

            DefaultParagraphStyle = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements.OfType<Style>()
                .SingleOrDefault(x => x.Type?.Value == StyleValues.Paragraph && (x.Default?.Value ?? false));
        }

        AllParagraphs = Document.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().ToList();
        
        FieldStackTracker.Run(document.MainDocumentPart!.Document!);

        if (Document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering is { } numbering)
        {
            foreach (var inst in numbering.ChildElements.OfType<NumberingInstance>())
            {
                var tool = new NumberingPropertiesTool(this, inst);

                foreach (var p in AllParagraphs)
                {
                    var pTool = GetTool(p);
                    
                    if (pTool.NumberingId == inst.NumberID?.Value)
                    {
                        tool.Paragraphs.Add(p);
                        pTool.OfNumbering = tool;
                    }
                }
                
                _numberings.Add(tool);
            }
        }
        
        StructuralElement? currentElement = null;
        ParagraphPropertiesTool? currentHeading1 = null;
        List<Paragraph> currentSection = [];
        
        foreach (var p in AllParagraphs)
        {
            var tool = GetTool(p);

            if (tool.StructuralElementHeader != null)
            {
                currentElement = tool.StructuralElementHeader;
            }

            if (tool.OutlineLevel == 0)
            {
                currentHeading1 = tool;
            }

            tool.OfStructuralElement = currentElement;
            tool.AssociatedHeading1 = currentHeading1;
            
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

        int i = 1;
        foreach (var p in AllParagraphs)
        {
            var tool = GetTool(p);
            
            if (tool is not {Class: ParagraphClass.Heading, OutlineLevel: 0}) continue;

            tool.HeadingNumber = i.ToString();
            i += 1;
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

    public Style? GetStyle(StyleValues styleType, string styleId)
    {
        return _styles.GetValueOrDefault((styleType, styleId));
    }

    public T? FollowStyleChain<T>(StyleValues styleType, string? styleId, Func<Style, T?> getter)
    {
        while (styleId != null)
        {
            var style = GetStyle(styleType, styleId);

            if (style == null) return default;

            var result = getter(style);
            if (result != null) return result;

            styleId = style.BasedOn?.Val?.Value;
        }

        return default;
    }

    public bool SniffStyleName(StyleValues styleType, string? styleId, string name)
    {
        while (styleId != null)
        {
            var style = GetStyle(styleType, styleId);

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

    private int _commentId = 0;
    
    public void WriteComment(LintMessage msg, DiagnosticTranslationsFile translations)
    {
        string id = (_commentId++).ToString();
        
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

        var translation = translations.Translate(msg.Id, msg.Parameters ?? new());
        c.Append(translation);

        commentsPart.Comments.AppendChild(c);
        commentsPart.Comments.Save();
    }

    public IReadOnlyList<SectionPropertiesTool> AllSections => _sections;
    public IEnumerable<Paragraph> AllParagraphs { get; }
}