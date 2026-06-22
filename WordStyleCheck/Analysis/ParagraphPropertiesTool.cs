using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace WordStyleCheck.Analysis;

public class ParagraphPropertiesTool : SupportsFeatures<ParagraphPropertiesTool>, IBlockLevelPropertiesTool
{
    public DocumentAnalysisContext Context { get; }
    public Paragraph Paragraph { get; }
    OpenXmlElement IBlockLevelPropertiesTool.Element => Paragraph;

    internal ParagraphPropertiesTool(DocumentAnalysisContext ctx, Paragraph paragraph)
    {
        Context = ctx;
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

        RunStyleId = Context.FollowStyleChain(StyleValues.Paragraph, styleId, x => x.LinkedStyle?.Val?.Value);
        
        OutlineLevel = FollowPropertyChain(
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value
        );

        if (OutlineLevel == 9) OutlineLevel = null;
        
        ProbablyHeading = OutlineLevel != null || Context.SniffStyleName(StyleValues.Paragraph, styleId, "Heading");
        PossiblyPartOfList = Context.SniffStyleName(StyleValues.Paragraph, styleId, "List");

        if (!IsTableOfContents && ContainingTableCell == null && NumberingId == null)
        {
            CaptionData = CaptionClassifierData.Classify(this, false);
        }

        EquationData = EquationClassifierData.Classify(this);
        
        IsEmptyOrDrawing = !Utils.DirectRunChildren(paragraph).SelectMany(x => x.ChildElements).Any(x => x is Text text && !string.IsNullOrWhiteSpace(text.Text));
        IsEmptyOrWhitespace = IsEmptyOrDrawing && !paragraph.Descendants().Any(x => x is Drawing or OfficeMath or DocumentFormat.OpenXml.Math.Paragraph);

        if (Utils.AscendToAnscestor<Table>(Paragraph) is { } table)
        {
            ContainingTable = ctx.GetTool(table);
        }
    }
    
    public string Contents { get; private set; }

    public List<RunPropertiesTool> Runs => Utils.DirectRunChildren(Paragraph).Select(x => Context.GetTool(x, this)).ToList();
    
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
    
    public int? RightIndent =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.Right,
            x => x.Indentation?.Right,
            x => x.Indentation?.Right
        )?.Value)
        ??
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.End,
            x => x.Indentation?.End,
            x => x.Indentation?.End
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
                Context.GetTool(p).Style?.StyleId == Style?.StyleId)
                return 0;

            if (FollowPropertyChain(
                    x => x.SpacingBetweenLines?.BeforeAutoSpacing?.Value,
                    x => x.SpacingBetweenLines?.BeforeAutoSpacing?.Value,
                    x => x.SpacingBetweenLines?.BeforeAutoSpacing?.Value
                ) ?? false)
                return 1337;
            
            int? beforeLines = FollowPropertyChain(
                x => x.SpacingBetweenLines?.BeforeLines?.Value,
                x => x.SpacingBetweenLines?.BeforeLines?.Value,
                x => x.SpacingBetweenLines?.BeforeLines?.Value
            );

            if (beforeLines != null) return beforeLines.Value; // TODO: make this good.
            
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

    public LineSpacingRuleValues LineSpacingRule =>
        FollowPropertyChain(
            x => x.SpacingBetweenLines?.LineRule?.Value,
            x => x.SpacingBetweenLines?.LineRule?.Value,
            x => x.SpacingBetweenLines?.LineRule?.Value
        ) ?? LineSpacingRuleValues.Auto;
    
    public int? ActualAfterSpacing
    {
        get
        {
            if (ContextualSpacing && this.Paragraph.NextSibling() is Paragraph p &&
                Context.GetTool(p).Style?.StyleId == Style?.StyleId)
                return 0;

            if (FollowPropertyChain(
                    x => x.SpacingBetweenLines?.AfterAutoSpacing?.Value,
                    x => x.SpacingBetweenLines?.AfterAutoSpacing?.Value,
                    x => x.SpacingBetweenLines?.AfterAutoSpacing?.Value
                ) ?? false)
                return 1337;
            
            int? afterLines = FollowPropertyChain(
                x => x.SpacingBetweenLines?.AfterLines?.Value,
                x => x.SpacingBetweenLines?.AfterLines?.Value,
                x => x.SpacingBetweenLines?.AfterLines?.Value
            );

            if (afterLines != null) return afterLines.Value; // TODO: make this good.
            
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

    public TablePropertiesTool? ContainingTable { get; }

    public TableRow? ContainingTableRow => Utils.AscendToAnscestor<TableRow>(Paragraph);

    public TableCell? ContainingTableCell => Utils.AscendToAnscestor<TableCell>(Paragraph);
    
    public TextBoxContent? ContainingTextBox => Utils.AscendToAnscestor<TextBoxContent>(Paragraph);
    
    public SdtBlock? ContainingSdtBlock => Utils.AscendToAnscestor<SdtBlock>(Paragraph);

    public bool IsTableOfContents
    {
        get
        {
            if (Context.GetContextFor(Paragraph).Any(x => x.InstrText != null && x.InstrText.Contains("TOC"))) return true;
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

    public bool MaybeParagraphContinuation
    {
        get
        {
            switch (Paragraph.PreviousSibling())
            {
                case Paragraph p when Context.GetTool(p).EquationData != null:
                case Table t when Context.GetTool(t).EquationData != null:
                    return Contents.Length > 0 && char.IsLower(Contents[0]);
                default:
                    return false;
            }
        }
    }

    public void ReloadContents()
    {
        StringBuilder contents = new();
        foreach (var r in Utils.DirectRunChildren(Paragraph))
        {
            var rTool = Context.GetTool(r, this);
            rTool.ReloadContents();
            
            contents.Append(rTool.Contents);
        }
        Contents = Utils.StripJunk(contents.ToString());        
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
            var result = Context.FollowStyleChain(StyleValues.Paragraph, Style.StyleId, x => x.StyleParagraphProperties != null ? styleGetter(x.StyleParagraphProperties) : default);
            if (result != null)
                return result;
        }

        if (Context.Document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle != null)
        {
            var result = baseStyleGetter(Context.Document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults
                ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle!);
            if (result != null)
                return result;
        }

        return default;
    }
}