using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public record ParagraphPropertiesTool(WordprocessingDocument Document, Paragraph Paragraph)
{
    public int? FirstLineIndent =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.Indentation?.FirstLine,
            x => x.Indentation?.FirstLine,
            x => x.Indentation?.FirstLine
        )?.Value);

    public int? OutlineLevel =>
        FollowPropertyChain(
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value,
            x => x.OutlineLevel?.Val?.Value
        );
    
    public int? BeforeSpacing =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.SpacingBetweenLines?.Before,
            x => x.SpacingBetweenLines?.Before,
            x => x.SpacingBetweenLines?.Before
        )?.Value);
    
    public int? LineSpacing =>
        Utils.ParseTwipsMeasure(FollowPropertyChain(
            x => x.SpacingBetweenLines?.Line,
            x => x.SpacingBetweenLines?.Line,
            x => x.SpacingBetweenLines?.Line
        )?.Value);

    public Style? Style
    {
        get
        {
            if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles == null) return null;
            
            if (Paragraph.ParagraphProperties?.ParagraphStyleId != null)
            {
                var style = Document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
                    .SingleOrDefault(x => x.StyleId?.Value == Paragraph.ParagraphProperties?.ParagraphStyleId.Val);

                return style;
            }

            {
                var style = Document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
                    .SingleOrDefault(x => x.Type?.Value == StyleValues.Paragraph && (x.Default?.Value ?? false));

                return style;
            }
        }
    }

    public string? RunStyleId
    {
        get
        {
            if (Document.MainDocumentPart?.StyleDefinitionsPart?.Styles == null) return null;
            
            string? FollowStyleChain(string? styleId)
            {
                if (styleId == null) return null;
                
                var style = Document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
                    .SingleOrDefault(x => x.StyleId?.Value == styleId);

                if (style == null) return null;

                if (style.LinkedStyle?.Val != null)
                {
                    return style.LinkedStyle.Val.Value;
                }

                if (style.BasedOn?.Val?.Value != null)
                {
                    return FollowStyleChain(style.BasedOn?.Val?.Value);
                }

                return null;
            }

            if (Paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } styleId)
                return FollowStyleChain(styleId);
            else
                return null;
        }
    }

    public TableCell? ContainingTableCell => Utils.AscendToAnscestor<TableCell>(Paragraph);

    public bool IsTableOfContents => FieldStackTracker.GetContextFor(Paragraph)
        .Any(x => x.InstrText != null && x.InstrText.Contains("TOC"));

    private T? FollowPropertyChain<T>(Func<ParagraphProperties, T?> getter, Func<StyleParagraphProperties, T?> styleGetter, Func<ParagraphPropertiesBaseStyle, T?> baseStyleGetter)
    {
        if (Paragraph.ParagraphProperties != null)
        {
            var result = getter(Paragraph.ParagraphProperties);
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

                if (style.StyleParagraphProperties != null)
                {
                    var result = styleGetter(style.StyleParagraphProperties);
                    if (result != null)
                        return result;
                }

                if (style.BasedOn?.Val?.Value != null)
                {
                    return FollowStyleChain(style.BasedOn?.Val?.Value);
                }

                return default;
            }
            
            if (Style?.StyleId != null)
            {
                var result = FollowStyleChain(Style.StyleId);
                if (result != null)
                    return result;
            }

            if (Document.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle != null)
            {
                var result = baseStyleGetter(Document.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults
                    ?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle!);
                if (result != null)
                    return result;
            }
        }

        return default;
    }
}