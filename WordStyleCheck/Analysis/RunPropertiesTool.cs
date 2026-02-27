using System.Diagnostics;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public record RunPropertiesTool
{
    private readonly DocumentAnalysisContext _ctx;
    public Run Run { get; }
 
    public RunPropertiesTool(DocumentAnalysisContext ctx, Run Run)
    {
        this._ctx = ctx;
        this.Run = Run;
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
    
    public Paragraph? ContainingParagraph => Utils.AscendToAnscestor<Paragraph>(Run);

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

        if (ContainingParagraph != null)
        {
            ParagraphPropertiesTool pTool = _ctx.GetTool(ContainingParagraph);

            // TODO: figure out in which order these two should go... or whether we should process linked styles at all...
            
            if (pTool.RunStyleId is {} rStyleId)
            {
                var result = FollowRunStyleChain(StyleValues.Character, rStyleId);
                if (result != null)
                    return result;
            }
            
            if (pTool.Style?.StyleId?.Value is { } pStyleId)
            {
                var result = FollowRunStyleChain(StyleValues.Paragraph, pStyleId);
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