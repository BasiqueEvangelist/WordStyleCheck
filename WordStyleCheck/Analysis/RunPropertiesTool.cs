using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public record RunPropertiesTool
{
    private readonly DocumentAnalysisContext _ctx;
    private readonly ParagraphPropertiesTool _parent;
    public Run Run { get; }
 
    public RunPropertiesTool(DocumentAnalysisContext ctx, ParagraphPropertiesTool parent, Run Run)
    {
        this._ctx = ctx;
        _parent = parent;
        this.Run = Run;
        
        StringBuilder contents = new();

        foreach (var t in Run.ChildElements.OfType<Text>())
        {
            contents.Append(t.Text);
        }

        Contents = contents.ToString();

        if (Caps)
        {
            Contents = Contents.ToUpper();
        }
    }

    public string? AsciiFont => FollowPropertyChain(
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value
    );
    
    public int? FontSize => Utils.ParseHpsMeasure(FollowPropertyChain(
        x => x.FontSize?.Val?.Value,
        x => x.FontSize?.Val?.Value,
        x => x.FontSize?.Val?.Value
    ));

    public bool Bold => FollowPropertyChain(
        x => Utils.ConvertOnOffType(x.Bold),
        x => Utils.ConvertOnOffType(x.Bold),
        x => Utils.ConvertOnOffType(x.Bold)
    ) ?? false;
    
    public bool Caps => FollowPropertyChain(
        x => Utils.ConvertOnOffType(x.Caps),
        x => Utils.ConvertOnOffType(x.Caps),
        x => Utils.ConvertOnOffType(x.Caps)
    ) ?? false;

    public string? Color => FollowPropertyChain(
        x => x.Color?.Val?.Value,
        x => x.Color?.Val?.Value,
        x => x.Color?.Val?.Value
    );

    public HighlightColorValues Highlight => Run.RunProperties?.Highlight?.Val?.Value ?? HighlightColorValues.None;

    public Paragraph ContainingParagraph => _parent.Paragraph;

    public bool IsHyperlink => Utils.AscendToAnscestor<Hyperlink>(Run) != null ||
                               _ctx.GetContextFor(Run).Any(x => x.InstrText?.Contains("\\h") ?? false);
    
    public string Contents { get; }

    private T? FollowPropertyChain<T>(Func<RunProperties, T?> getter, Func<StyleRunProperties, T?> styleGetter, Func<RunPropertiesBaseStyle, T?> baseStyleGetter)
    {
        if (Run.RunProperties != null)
        {
            var result = getter(Run.RunProperties);
            if (result != null)
                return result;
        }
        
        T? FollowRunStyleChain(StyleValues type, string? styleId)
        {
            if (styleId == null) return default;
            
            var style = _ctx.GetStyle(type, styleId);

            if (style == null) return default;

            if (style.StyleRunProperties != null)
            {
                var result = styleGetter(style.StyleRunProperties);
                if (result != null)
                    return result;
            }

            if (style.BasedOn?.Val?.Value != null)
            {
                return FollowRunStyleChain(type, style.BasedOn?.Val?.Value);
            }

            return default;
        }

        {
            var styleId = Run.RunProperties?.RunStyle?.Val?.Value;

            var result = FollowRunStyleChain(StyleValues.Character, styleId);
            if (result != null)
                return result;
        }

        {
            // TODO: figure out in which order these two should go... or whether we should process linked styles at all...
            
            if (_parent.Style?.StyleId?.Value is { } pStyleId)
            {
                var result = FollowRunStyleChain(StyleValues.Paragraph, pStyleId);
                if (result != null)
                    return result;
            }
            
            if (_parent.RunStyleId is {} rStyleId)
            {
                var result = FollowRunStyleChain(StyleValues.Character, rStyleId);
                if (result != null)
                    return result;
            }
        }

        if (_ctx.Document?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle != null)
        {
            var result = baseStyleGetter(_ctx.Document.MainDocumentPart.StyleDefinitionsPart.Styles!.DocDefaults!
                .RunPropertiesDefault?.RunPropertiesBaseStyle!);
            if (result != null)
                return result;
        }

        return default;
    }
}