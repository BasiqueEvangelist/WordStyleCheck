using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class ParagraphPropertiesTool : SupportsFeatures<ParagraphPropertiesTool>
{
    private readonly DocumentAnalysisContext _ctx;
    public Paragraph Paragraph { get; }

    internal ParagraphPropertiesTool(DocumentAnalysisContext ctx, Paragraph paragraph)
    {
        _ctx = ctx;
        Paragraph = paragraph;

        StringBuilder contents = new();
        foreach (var r in Utils.DirectRunChildren(paragraph))
        {
            contents.Append(ctx.GetTool(r, this).Contents);
        }
        Contents = Utils.StripJunk(contents.ToString());
       
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        
        Style = new Func<Style?>(() =>
        {
            if (styleId != null)
            {
                return ctx.GetStyle(StyleValues.Paragraph, styleId);
            }

            return ctx.DefaultParagraphStyle;
        })();

        RunStyleId = _ctx.FollowStyleChain(StyleValues.Paragraph, styleId, x => x.LinkedStyle?.Val?.Value);
        
        OutlineLevel = FollowPropertyChain(
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value
        );

        if (OutlineLevel == 9) OutlineLevel = null;
        
        ProbablyHeading = OutlineLevel != null || _ctx.SniffStyleName(StyleValues.Paragraph, styleId, "Heading");
        PossiblyPartOfList = _ctx.SniffStyleName(StyleValues.Paragraph, styleId, "List");

        if (!IsTableOfContents && ContainingTableCell == null && NumberingId == null)
        {
            CaptionData = CaptionClassifierData.Classify(this, false);
        }

        EquationData = EquationClassifierData.Classify(this);
        
        IsEmptyOrDrawing = !Utils.DirectRunChildren(paragraph).SelectMany(x => x.ChildElements).Any(x => x is Text text && !string.IsNullOrWhiteSpace(text.Text));
        IsEmptyOrWhitespace = IsEmptyOrDrawing && !paragraph.Descendants().Any(x => x is Drawing);
    }
    
    public string Contents { get; }

    public List<RunPropertiesTool> Runs => Utils.DirectRunChildren(Paragraph).Select(x => _ctx.GetTool(x, this)).ToList();
    
    public int? FirstLineIndent
    {
        get
        {
            return FollowPropertyChain(
                x => Utils.ParseTwipsMeasure(x.Indentation?.FirstLine) ?? -Utils.ParseTwipsMeasure(x.Indentation?.Hanging),
                x => Utils.ParseTwipsMeasure(x.Indentation?.FirstLine) ?? -Utils.ParseTwipsMeasure(x.Indentation?.Hanging),
                x => Utils.ParseTwipsMeasure(x.Indentation?.FirstLine) ?? -Utils.ParseTwipsMeasure(x.Indentation?.Hanging)
            );
        }
    }

    public int? LeftIndent =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.Left,
            x => x.Indentation?.Left,
            x => x.Indentation?.Left
        )?.Value)
        ??
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.Start,
            x => x.Indentation?.Start,
            x => x.Indentation?.Start
        )?.Value);

    public int? OutlineLevel { get; }
    
    public bool ContextualSpacing => FollowPropertyChain(
        x => Utils.ConvertOnOffType(x.ContextualSpacing),
        x => Utils.ConvertOnOffType(x.ContextualSpacing),
        x => Utils.ConvertOnOffType(x.ContextualSpacing)
    ) ?? false;
    
    public bool PageBreakBefore => FollowPropertyChain(
        x => Utils.ConvertOnOffType(x.PageBreakBefore),
        x => Utils.ConvertOnOffType(x.PageBreakBefore),
        x => Utils.ConvertOnOffType(x.PageBreakBefore)
    ) ?? false;
    
    public int? ActualBeforeSpacing
    {
        get
        {
            if (ContextualSpacing && this.Paragraph.PreviousSibling() is Paragraph p &&
                _ctx.GetTool(p).Style?.StyleId == Style?.StyleId)
                return 0;
            
            return Utils.ParseTwipsMeasure(FollowPropertyChain(
                x => x.SpacingBetweenLines?.Before,
                x => x.SpacingBetweenLines?.Before,
                x => x.SpacingBetweenLines?.Before
            )?.Value);
        }
    }

    public int? LineSpacing =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.SpacingBetweenLines?.Line,
            x => x.SpacingBetweenLines?.Line,
            x => x.SpacingBetweenLines?.Line
        )?.Value);
    
    public int? ActualAfterSpacing
    {
        get
        {
            if (ContextualSpacing && this.Paragraph.NextSibling() is Paragraph p &&
                _ctx.GetTool(p).Style?.StyleId == Style?.StyleId)
                return 0;
            
            return Utils.ParseTwipsMeasure(FollowPropertyChain(
                x => x.SpacingBetweenLines?.After,
                x => x.SpacingBetweenLines?.After,
                x => x.SpacingBetweenLines?.After
            )?.Value);
        }
    }
    
    public JustificationValues? Justification =>
        FollowPropertyChain(
            x => x.Justification?.Val?.Value,
            x => x.Justification?.Val?.Value,
            x => x.Justification?.Val?.Value
        );

    public Style? Style { get; }

    public string? RunStyleId { get; }

    public TableRow? ContainingTableRow => Utils.AscendToAnscestor<TableRow>(Paragraph);

    public TableCell? ContainingTableCell => Utils.AscendToAnscestor<TableCell>(Paragraph);
    
    public TextBoxContent? ContainingTextBox => Utils.AscendToAnscestor<TextBoxContent>(Paragraph);
    
    public SdtBlock? ContainingSdtBlock => Utils.AscendToAnscestor<SdtBlock>(Paragraph);

    public bool IsTableOfContents
    {
        get
        {
            if (_ctx.GetContextFor(Paragraph).Any(x => x.InstrText != null && x.InstrText.Contains("TOC"))) return true;
            if (ContainingSdtBlock is {SdtProperties: {} props} && props.Descendants<DocPartGallery>().Any(x => x.Val?.Value == "Table of Contents"))
                return true;
            
            return false;
        }
    }

    public bool ProbablyHeading { get; }

    public bool ProbablyCodeListing { get; internal set; }

    public bool ProbablyTableColumnHeader { get; internal set; }

    public int? NumberingId
    {
        get
        {
            int? val = FollowPropertyChain(
                x => x.NumberingProperties?.NumberingId?.Val?.Value,
                x => x.NumberingProperties?.NumberingId?.Val?.Value,
                x => x.NumberingProperties?.NumberingId?.Val?.Value
            );

            if (val == 0) val = null;

            return val;
        }
    }

    public int NumberingLevel => FollowPropertyChain(
        x => x.NumberingProperties?.NumberingLevelReference?.Val?.Value,
        x => x.NumberingProperties?.NumberingLevelReference?.Val?.Value,
        x => x.NumberingProperties?.NumberingLevelReference?.Val?.Value
    ) ?? 0;
    
    public INumbering? OfNumbering { get; internal set; }
    
    public bool PossiblyPartOfList { get; }

    public ParagraphPropertiesTool? AssociatedHeading1 { get; internal set; }
    
    public CaptionClassifierData? CaptionData { get; internal set; }
    
    public HeadingClassifierData? HeadingData { get; internal set; }
    
    public EquationClassifierData? EquationData { get; }
    
    public bool IsEmptyOrWhitespace { get; }
    
    public bool IsEmptyOrDrawing { get; }

    public bool IsIgnored { get; set; }
    
    private T? FollowPropertyChain<T>(Func<ParagraphProperties, T?> getter, Func<StyleParagraphProperties, T?> styleGetter, Func<ParagraphPropertiesBaseStyle, T?> baseStyleGetter)
    {
        if (Paragraph.ParagraphProperties != null)
        {
            var result = getter(Paragraph.ParagraphProperties);
            if (result != null)
                return result;
        }
        
        if (Style?.StyleId != null)
        {
            var result = _ctx.FollowStyleChain(StyleValues.Paragraph, Style.StyleId, x => x.StyleParagraphProperties != null ? styleGetter(x.StyleParagraphProperties) : default);
            if (result != null)
                return result;
        }

        if (_ctx.Document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle != null)
        {
            var result = baseStyleGetter(_ctx.Document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
                ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle!);
            if (result != null)
                return result;
        }

        return default;
    }
}