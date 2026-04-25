using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Gost7_32;

public class Gost7_32Profile : IProfile
{
    public List<IClassifier> Classifiers { get; } =
    [
        new AttachGostDataClassifier(),
        new GostStructuralElementClassifier(),
        new TextAssociationClassifier()
    ];
    public List<ILint> Lints { get; } = 
    [
        new PageSizeLint(),
        new PageMarginsLint(new PageMargins(1134, 1134, 1701, 851, 709, 709, 0)),
        new TocReferencesLint(ShouldBeInToc),
        new HandmadeListLint(),
        new HandmadePageBreakLint(),
        new NeedlessParagraphLint(x => x.GetFeature(GostParagraphData.Key)!.Class == GostParagraphClass.BodyText),
        new ForcePageBreakBeforeLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.HeadingData.Level: 1} or {Class: GostParagraphClass.StructuralElementHeader}, "NeedsPageBreakBeforeHeader"),
        new ForceJustificationLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.StructuralElementHeader}, [JustificationValues.Center], "StructuralElementHeaderNotCentered"),
        new ForceJustificationLint(x => x is {CaptionData.Type: CaptionType.Table}, [JustificationValues.Left, JustificationValues.Both], "TableCaptionNotLeftAligned"),
        new ParagraphIndentLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.BodyText, Inner.OfNumbering: null, Inner.PossiblyPartOfList: false}, 709, 0, "IncorrectBodyTextFirstLineIndent", "IncorrectBodyTextLeftIndent"),
        new ParagraphIndentLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 0}, 709, 0, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 1}, -709, 1418, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphIndentLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 2}, -851, 1560, "IncorrectHeadingFirstLineIndent", "IncorrectHeadingLeftIndent"),
        new ParagraphLineSpacingLint(
            // TODO: enforce this for numberings.
            // TODO: enforce this for table cell content, headers, captions.
            x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.BodyText, Inner.OfNumbering: null, Inner.IsEmptyOrDrawing: false},
            360,
            "IncorrectTextLineSpacing"
        ),
        new ParagraphLineSpacingLint(
            // TODO: enforce this for numberings.
            // TODO: enforce this for table cell content, headers, captions.
            x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Caption},
            240,
            "IncorrectCaptionLineSpacing"
        ),
        new InterParagraphSpacingLint(
            [
                new(
                    x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 0},
                    0,
                    18 * 20
                ),
                new(
                    x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 1},
                    24 * 20,
                    12 * 20
                ),
                new(
                    x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading, Inner.OutlineLevel: 2},
                    12 * 20,
                    6 * 20
                ),
                new(
                    x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.BodyText, Inner.OfNumbering: null, Inner.IsIgnored: false},
                    6 * 20,
                    6 * 20,
                    contextualSpacing: true
                )
            ],
            "IncorrectInterParagraphSpacing"
        ),
        new CorrectStructuralElementHeaderLint(),
        new WrongCaptionPositionLint(CaptionType.Table, false, "IncorrectTableCaptionPosition"),
        new WrongCaptionPositionLint(CaptionType.Figure, true, "IncorrectFigureCaptionPosition"),
        new WrongCaptionPositionLint(CaptionType.Listing, true, "IncorrectListingCaptionPosition"),
        new IncorrectCaptionTextLint(),
        new IncorrectCaptionedNumberingLint(x => x.GetFeature(GostParagraphData.Key)!.OfStructuralElement != GostStructuralElement.Appendix, CaptionType.Figure, "IncorrectFigureNumbering", "FigureNumberingMix"),
        new IncorrectCaptionedNumberingLint(x => x.GetFeature(GostParagraphData.Key)!.OfStructuralElement != GostStructuralElement.Appendix, CaptionType.Listing, "IncorrectListingNumbering", "ListingNumberingMix"),
        new IncorrectCaptionedNumberingLint(x => x.GetFeature(GostParagraphData.Key)!.OfStructuralElement != GostStructuralElement.Appendix, CaptionType.Table, "IncorrectTableNumbering", "TableNumberingMix"),
        new FigureTableNotReferencedLint(),
        new BibliographySourceNotReferencedLint(x => x.GetFeature(GostParagraphData.Key)!.OfStructuralElement == GostStructuralElement.Bibliography),
        new IncorrectHeadingTextLint(),
        // TODO: make this lint configurable.
        // new NotEnoughSourcesLint(40, "NotEnoughSources", "NoBibliography"),
        // new IncorrectOutlineLevelLint(x => x is { Class: ParagraphClass.BodyText }, _ => null, "BodyTextInToC"),
        // new IncorrectOutlineLevelLint(x => x is { HeadingData.Level: < 4 }, x => x.HeadingData!.Level - 1, "IncorrectHeaderOutlineLevel"),
        // new IncorrectOutlineLevelLint(x => x is { HeadingData.Level: 4 }, x => null, "SubPointsInToC"),
        new HeadingOutlineLevelLint(x => x is {HeadingData: not null} || x.GetFeature(GostParagraphData.Key) is {OfStructuralElement: (GostStructuralElement.Introduction or GostStructuralElement.Conclusion or GostStructuralElement.Bibliography or GostStructuralElement.Appendix)}),
        new IncorrectHeadingNumberingLint(),
        // TODO: make text font lint configurable.
        // new TextFontLint(),
        new FontSizeLint(x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.Heading or GostParagraphClass.BodyText}, 24, false, "IncorrectFontSize"),
        new ForceBoldLint(true, x => x.GetFeature(GostParagraphData.Key) is { Class: GostParagraphClass.Heading, Inner.OutlineLevel: null or < 2 } or {Class: GostParagraphClass.StructuralElementHeader}, "HeadingNotBold"),
        new ForceBoldLint(false, x => x is { OutlineLevel: >= 2 }, "SubSubHeadingBold"),
        new ForceBoldLint(false, x => x.GetFeature(GostParagraphData.Key) is {Class: GostParagraphClass.BodyText}, "BodyTextBold"),
        new TextColorLint()
    ];
    
    private static bool ShouldBeInToc(ParagraphPropertiesTool tool)
    {
        var data = tool.GetFeature(GostParagraphData.Key)!;
        
        if (data.OfStructuralElement == GostStructuralElement.Appendix &&
            data.StructuralElementHeader != GostStructuralElement.Appendix)
            return false;

        if (tool is { HeadingData.Level: < 4 } or { HeadingData.IsConclusion: true })
            return true;

        if (data is
            {
                StructuralElementHeader: GostStructuralElement.Introduction or GostStructuralElement.Conclusion
                or GostStructuralElement.Bibliography or GostStructuralElement.Appendix
            })
            return true;
        
        return false;
    }
}