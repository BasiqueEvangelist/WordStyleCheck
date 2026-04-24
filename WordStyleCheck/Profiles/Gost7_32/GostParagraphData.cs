using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Gost7_32;

public class GostParagraphData
{
    public static readonly FeatureKey<GostParagraphData, ParagraphPropertiesTool> Key = new();

    public ParagraphPropertiesTool Inner { get; }

    public GostParagraphData(ParagraphPropertiesTool tool)
    {
        Inner = tool;
    }
    
    public GostStructuralElement? StructuralElementHeader { get; set; }

    public GostStructuralElement? OfStructuralElement { get; set; }
    
    public GostParagraphClass Class
    {
        get
        {
            if (Inner.ProbablyCodeListing) return GostParagraphClass.CodeListing;
            if (StructuralElementHeader != null) return GostParagraphClass.StructuralElementHeader;
            if (Inner.IsTableOfContents) return GostParagraphClass.TableOfContents;
            if (Inner.CaptionData != null) return GostParagraphClass.Caption;
            if (Inner.EquationData != null) return GostParagraphClass.DisplayEquation;
            if (Inner.ContainingTextBox != null) return GostParagraphClass.InsideDrawing;
            if (Inner.ProbablyTableColumnHeader) return GostParagraphClass.TableColumnHeader;
            if (Inner.ContainingTableCell != null) return GostParagraphClass.TableContent;
            if (Inner.ProbablyHeading || Inner.HeadingData != null) return GostParagraphClass.Heading;
            if (OfStructuralElement == GostStructuralElement.Appendix) return GostParagraphClass.AppendixText;
            if (OfStructuralElement == GostStructuralElement.Bibliography)
                return Inner.OfNumbering != null
                    ? GostParagraphClass.BibliographySource
                    : GostParagraphClass.BibliographyMisc;

            return GostParagraphClass.BodyText;
        }
    }
}

public enum GostParagraphClass
{
    BodyText,
    StructuralElementHeader,
    Heading,
    TableColumnHeader,
    TableContent,
    Caption,
    TableOfContents,
    CodeListing,
    InsideDrawing,
    DisplayEquation,
    AppendixText,
    BibliographySource,
    BibliographyMisc,
}