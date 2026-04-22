using DocumentFormat.OpenXml;
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
    private readonly Dictionary<int, NumberingPropertiesTool> _numberings = [];

    private readonly HashSet<string> _existingComments = [];

    private readonly Dictionary<OpenXmlElement, List<FieldStackTracker.FieldStackEntry>> _fieldStacks;

    public Style? DefaultParagraphStyle { get; }

    public List<HandmadeListClassifier.SniffedListData> HandmadeLists { get; }
    
    public Dictionary<string, BookmarkStart> BookmarkStarts { get; }

    public DocumentAnalysisContext(WordprocessingDocument document, List<IClassifier> classifiers)
    {
        Document = document;

        using (new LoudStopwatch("Locating styles"))
        {
            if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
            {
                _styles = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements.OfType<Style>()
                    .ToDictionary(x => (x.Type!.Value!, x.StyleId!.Value!), x => x);

                DefaultParagraphStyle = Document.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements
                    .OfType<Style>()
                    .SingleOrDefault(x => x.Type?.Value == StyleValues.Paragraph && (x.Default?.Value ?? false));
            }
        }

        AllParagraphs = Document.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().ToList();
        
        _fieldStacks = FieldStackTracker.RunTracker(document.MainDocumentPart!.Document!);

        using (new LoudStopwatch("Generating ParagraphPropertiesTool objects"))
            foreach (var p in AllParagraphs)
            {
                GetTool(p);
            }

        using (new LoudStopwatch("Assigning numberings"))
        {
            if (Document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering is { } numbering)
            {
                foreach (var inst in numbering.ChildElements.OfType<NumberingInstance>())
                {
                    var tool = new NumberingPropertiesTool(this, inst);
                    _numberings[inst.NumberID!.Value!] = tool;
                }

                foreach (var p in AllParagraphs)
                {
                    var pTool = GetTool(p);

                    if (pTool.NumberingId is { } numId && _numberings.TryGetValue(numId, out var tool))
                    {
                        tool.Paragraphs.Add(p);
                        pTool.OfNumbering = tool;
                    }
                }
            }
        }

        StructuralElement? currentElement = null;
        List<Paragraph> currentSection = [];
        
        using (new LoudStopwatch("Assigning OfStructuralElement and page style sections"))
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
        
        if (Document.MainDocumentPart.WordprocessingCommentsPart?.Comments is { } comments)
        {
            _existingComments = comments.ChildElements.OfType<Comment>().Select(x => x.Id?.Value ?? "").Distinct().ToHashSet();
        }

        HashSet<OpenXmlElement> alreadyReferenced = [];
        foreach (var p in AllParagraphs)
        {
            if (GetTool(p).CaptionData is not {TargetedElement: {} targeted}) continue;

            alreadyReferenced.Add(targeted);
        }

        using (new LoudStopwatch("Second caption classification pass"))
            foreach (var p in AllParagraphs)
            {
                var tool = GetTool(p);

                if (tool.Class != ParagraphClass.BodyText) continue;

                var caption = CaptionClassifierData.Classify(tool, true);
                
                if (!caption.HasValue) continue;
                
                if (caption.Value.TargetedElement != null && alreadyReferenced.Contains(caption.Value.TargetedElement))
                    continue;

                tool.CaptionData = caption;
            }

        using (new LoudStopwatch("HandmadeListClassifier.Classify"))
            HandmadeLists = HandmadeListClassifier.Classify(this);

        using (new LoudStopwatch("HeadingClassifierData.Classify and classifying code listings"))
            foreach (var p in AllParagraphs)
            {
                var tool = GetTool(p);
                tool.HeadingData = HeadingClassifierData.Classify(tool);

                bool isListing = false;
                
                foreach (var run in Utils.DirectRunChildren(p))
                {
                    RunPropertiesTool runTool = GetTool(run);
                
                    if (string.IsNullOrWhiteSpace(Utils.CollectText(run))) continue;

                    if (!Utils.IsMonospaceFont(runTool.AsciiFont ?? "<?>"))
                    {
                        isListing = false;
                        break;
                    }

                    isListing = true;
                }

                tool.ProbablyCodeListing = isListing;
            }
        
        HashSet<OpenXmlElement> continuationTables = [];
        using (new LoudStopwatch("Finding Continuation Tables"))
            foreach (var p in AllParagraphs)
            {
                if (GetTool(p) is { CaptionData: { IsContinuation: true, Type: CaptionType.Table, TargetedElement: { } targeted } })
                {
                    continuationTables.Add(targeted);
                }
            }

        using (new LoudStopwatch("Assigning ProbableTableColumnHeader"))
            foreach (var p in AllParagraphs)
            {
                var tool = GetTool(p);

                if (tool.ContainingTableRow is not { } tr) continue;

                var table = (Table?)tr.Parent;

                if (table == null || continuationTables.Contains(table)) continue;

                int rowIndex = table.ChildElements.ToList().IndexOf(tr);

                if (rowIndex == 0)
                {
                    tool.ProbablyTableColumnHeader = true;
                }
            }

        using (new LoudStopwatch("Locating Bookmark Starts"))
            BookmarkStarts = Document.MainDocumentPart.Document.Body.Descendants<BookmarkStart>()
                .DistinctBy(x => x.Name!.Value!)
                .ToDictionary(x => x.Name!.Value!, x => x);

        foreach (var classifier in classifiers)
        {
            classifier.Classify(this);
        }
    }

    private static readonly List<FieldStackTracker.FieldStackEntry> _emptyList = [];
    public List<FieldStackTracker.FieldStackEntry> GetContextFor(OpenXmlElement el) => _fieldStacks.GetValueOrDefault(el, _emptyList);
    
    public ParagraphPropertiesTool GetTool(Paragraph p)
    {
        if (_paragraphTools.TryGetValue(p, out var tool)) return tool;
        
        tool = new ParagraphPropertiesTool(this, p);
        _paragraphTools[p] = tool;

        return tool;
    }
    
    public RunPropertiesTool GetTool(Run r, ParagraphPropertiesTool? parent = null)
    {
        if (_runTools.TryGetValue(r, out var tool)) return tool;

        if (parent == null) parent = GetTool(Utils.AscendToAnscestor<Paragraph>(r)!);
        
        tool = new RunPropertiesTool(this, parent, r);
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
    
    public Comment WriteComment(LintMessage msg, XmlTranslationsFile translations)
    {
        string id = (_commentId++).ToString();

        while (_existingComments.Contains(id))
        {
            id = (_commentId++).ToString();
        }
        
        msg.Context.WriteCommentReference(id);

        var commentsPart = Document.MainDocumentPart!.WordprocessingCommentsPart ??
            Document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();

        commentsPart.Comments ??= new Comments();

        Comment c = new Comment()
        {
            Id = id,
            Author = "WordStyleCheck",
            Initials = "WSC",
            Date = DateTime.Now
        };

        var translation = translations.Translate(msg.Id, msg.Parameters ?? new(), this);
        c.Append(translation);

        commentsPart.Comments.AppendChild(c);

        return c;
    }

    public string AllocateHyperlinkRelationship(Uri url)
    {
        var existing = Document.MainDocumentPart!.HyperlinkRelationships.FirstOrDefault(x => x.Uri == url);
        if (existing != null) return existing.Id;

        return Document.MainDocumentPart.AddHyperlinkRelationship(url, true).Id;
    }

    public IReadOnlyList<SectionPropertiesTool> AllSections => _sections;
    public IEnumerable<Paragraph> AllParagraphs { get; }
}