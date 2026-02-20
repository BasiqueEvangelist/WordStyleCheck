using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public record RunPropertiesTool(WordprocessingDocument Document, Run Run)
{
    public string? AsciiFont => FollowPropertyChain(
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value
    );
    
    private T? FollowPropertyChain<T>(Func<RunProperties, T?> getter, Func<StyleRunProperties, T?> styleGetter, Func<RunPropertiesBaseStyle, T?> baseStyleGetter)
    {
        if (Run.RunProperties != null)
        {
            var result = getter(Run.RunProperties);
            if (result != null)
                return result;
        }
        
        if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
        {
            T? FollowStyleChain(string? styleId)
            {
                if (styleId == null) return default;
                
                var style = Document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
                    .SingleOrDefault(x => x.StyleId?.Value == styleId);

                if (style == null) return default;

                if (style.StyleRunProperties != null)
                {
                    var result = styleGetter(style.StyleRunProperties);
                    if (result != null)
                        return result;
                }

                if (style.BasedOn?.Val?.Value != null)
                {
                    FollowStyleChain(style.BasedOn?.Val?.Value);
                }

                return default;
            }
            
            var styleId = Run.RunProperties?.RunStyle?.Val?.Value;

            {
                var result = FollowStyleChain(styleId);
                if (result != null)
                    return result;
            }

            if (Document.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle != null)
            {
                var result = baseStyleGetter(Document.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults
                    ?.RunPropertiesDefault?.RunPropertiesBaseStyle!);
                if (result != null)
                    return result;
            }
        }

        return default;
    }
}