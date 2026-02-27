using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public record ParagraphPropertiesTool
{
    private readonly DocumentAnalysisContext _ctx;
    public Paragraph Paragraph { get; init; }

    internal ParagraphPropertiesTool(DocumentAnalysisContext ctx, Paragraph Paragraph)
    {
        _ctx = ctx;
        this.Paragraph = Paragraph;
       
        string? styleId = Paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        
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
        
        ProbablyHeading = OutlineLevel != null || _ctx.SniffStyleName(StyleValues.Paragraph, styleId, "Heading");

        string? runFont = Paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<RunFonts>()?.Ascii?.Value;
        if (runFont == null && styleId != null)
        {
            runFont = _ctx.FollowStyleChain(StyleValues.Paragraph, styleId, x => x.StyleRunProperties?.RunFonts)?.Ascii;
        }

        if (runFont != null)
        {
            if (runFont.Contains("Code") || runFont == "Consolas" || runFont.Contains("Mono"))
                ProbablyCodeListing = true;
        }

        if (!IsTableOfContents)
        {
            StructuralElementHeader = StructuralElementHeaderClassifier.Classify(Paragraph);
            CaptionData = CaptionClassifierData.Classify(Paragraph);
        }
    }
    
    public int? FirstLineIndent
    {
        get
        {
            int? positive = Utils.ParseTwipsMeasure(FollowPropertyChain(
                x => x.Indentation?.FirstLine,
                x => x.Indentation?.FirstLine,
                x => x.Indentation?.FirstLine
            )?.Value);
            
            int? negative = Utils.ParseTwipsMeasure(FollowPropertyChain(
                x => x.Indentation?.Hanging,
                x => x.Indentation?.Hanging,
                x => x.Indentation?.Hanging
            )?.Value);

            return positive ?? -negative;
        }
    }

    public int? LeftIndent =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.Left,
            x => x.Indentation?.Left,
            x => x.Indentation?.Left
        )?.Value);

    public int? OutlineLevel { get; }
    
    public bool ContextualSpacing => FollowPropertyChain(
        x => Utils.ConvertOnOffType(x.ContextualSpacing),
        x => Utils.ConvertOnOffType(x.ContextualSpacing),
        x => Utils.ConvertOnOffType(x.ContextualSpacing)
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

    public TableCell? ContainingTableCell => Utils.AscendToAnscestor<TableCell>(Paragraph);

    public bool IsTableOfContents => FieldStackTracker.GetContextFor(Paragraph)
        .Any(x => x.InstrText != null && x.InstrText.Contains("TOC"));
    
    public bool ProbablyHeading { get; }

    public bool ProbablyCodeListing { get; }

    public int? NumberingId => FollowPropertyChain(
        x => x.NumberingProperties?.NumberingId?.Val?.Value,
        x => x.NumberingProperties?.NumberingId?.Val?.Value,
        x => x.NumberingProperties?.NumberingId?.Val?.Value
    );
    
    public INumbering? OfNumbering { get; internal set; }

    public StructuralElement? StructuralElementHeader { get; }
    
    public StructuralElement? OfStructuralElement { get; internal set; }
    
    public ParagraphPropertiesTool? AssociatedHeading1 { get; internal set; }
    
    public string? HeadingNumber { get; internal set; }
    
    public CaptionClassifierData? CaptionData { get; }
    
    public ParagraphClass Class
    {
        get
        {
            if (StructuralElementHeader != null) return ParagraphClass.StructuralElementHeader;
            if (IsTableOfContents) return ParagraphClass.TableOfContents;
            if (ContainingTableCell != null) return ParagraphClass.TableContent;
            if (CaptionData != null) return ParagraphClass.Caption;
            if (ProbablyHeading) return ParagraphClass.Heading;
            if (ProbablyCodeListing) return ParagraphClass.CodeListing;

            return ParagraphClass.BodyText;
        }
    }

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

public enum ParagraphClass
{
    BodyText,
    StructuralElementHeader,
    Heading, // TODO: Headings of different levels.
    TableContent,
    Caption,
    TableOfContents,
    CodeListing,
}