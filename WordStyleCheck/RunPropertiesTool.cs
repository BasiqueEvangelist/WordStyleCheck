using System.Diagnostics;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public record RunPropertiesTool
{
    private static readonly ConditionalWeakTable<Run, RunPropertiesTool> Cache = new();

    public static RunPropertiesTool Get(WordprocessingDocument document, Run run)
    {
        return Cache.GetValue(run, r => new RunPropertiesTool(document, r));
    }
 
    private RunPropertiesTool(WordprocessingDocument Document, Run Run)
    {
        this.Document = Document;
        this.Run = Run;
    }
    
    public WordprocessingDocument Document { get; }
    public Run Run { get; }

    public string? AsciiFont => FollowPropertyChain(
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value,
        x => x.RunFonts?.Ascii?.Value
    );
    
    public Paragraph? ContainingParagraph => Utils.AscendToAnscestor<Paragraph>(Run);

    private T? FollowPropertyChain<T>(Func<RunProperties, T?> getter, Func<StyleRunProperties, T?> styleGetter, Func<RunPropertiesBaseStyle, T?> baseStyleGetter)
    {
        if (Run.Descendants<Text>().Aggregate("", (a, b) => a + b.Text).Contains("Продолжение таблицы"))
        {
            // Debugger.Break();
        }
        
        if (Run.RunProperties != null)
        {
            var result = getter(Run.RunProperties);
            if (result != null)
                return result;
        }
        
        if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
        {
            T? FollowRunStyleChain(string? styleId)
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
                    return FollowRunStyleChain(style.BasedOn?.Val?.Value);
                }

                return default;
            }

            {
                var styleId = Run.RunProperties?.RunStyle?.Val?.Value;

                var result = FollowRunStyleChain(styleId);
                if (result != null)
                    return result;
            }

            if (ContainingParagraph != null)
            {
                ParagraphPropertiesTool pTool = ParagraphPropertiesTool.Get(Document, ContainingParagraph);

                // TODO: figure out in which order these two should go... or whether we should process linked styles at all...
                
                if (pTool.RunStyleId != null)
                {
                    var result = FollowRunStyleChain(pTool.RunStyleId);
                    if (result != null)
                        return result;
                }
                
                if (pTool.Style?.StyleId?.Value is { } pStyleId)
                {
                    var result = FollowRunStyleChain(pStyleId);
                    if (result != null)
                        return result;
                }
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

    public void Deconstruct(out WordprocessingDocument Document, out Run Run)
    {
        Document = this.Document;
        Run = this.Run;
    }
}