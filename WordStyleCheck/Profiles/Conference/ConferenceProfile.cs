using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Conference;

public class ConferenceProfile : IProfile
{
    public List<IClassifier> Classifiers { get; } =
    [
        new ConferenceParagraphData.Attacher(),
        new ConferencePartsClassifier()
    ];

    public List<ILint> Lints { get; } =
    [
        new AtLeastOneLint(IsOfClass(ConferenceParagraphClass.UniversalDecimalClassifier), "NoUdc", false),
        new AtLeastOneLint(IsOfClass(ConferenceParagraphClass.AuthorDetails), "NoAuthors", false),
        new AtLeastOneLint(IsOfClass(ConferenceParagraphClass.ThesisTitle), "NoThesisTitle", false),
        new AtLeastOneLint(IsOfClass(ConferenceParagraphClass.BibliographyHeader), "NoBibliography", true),
        new AtLeastOneLint(IsOfClass(ConferenceParagraphClass.Copyright), "NoCopyright", true),
        
        new PageSizeLint(false),
        new TextFontLint(),
        new FontSizeLint(
            IsBodyText,
            28,
            true,
            "IncorrectBodyTextFontSize"
        ),
        // TODO: links.
        new PageMarginsLint(new PageMargins(1134, 1134, 1134, 1134, 709, 709, 0)),
        new ForceJustificationLint(IsBodyText, [JustificationValues.Both], "BodyTextNotJustified"),
        new ParagraphLineSpacingLint(IsBodyText, 288, "IncorrectBodyTextLineSpacing"),
        new ParagraphIndentLint(IsBodyText, 566, 0, "IncorrectBodyTextFirstLineIndent", "IncorrectBodyTextLeftIndent"),
        new ForceCapsLint(IsOfClass(ConferenceParagraphClass.ThesisTitle), "ThesisTitleMustBeCaps"),
        new FontSizeLint(IsOfClass(ConferenceParagraphClass.ThesisTitle), 30, true, "IncorrectThesisTitleFontSize"),
        // TODO: force quote style.
        new ForceItalicLint(true, x => x.CaptionData is {Type: CaptionType.Figure}, "FigureCaptionNotItalic"),
        new FontSizeLint(IsOfClass(ConferenceParagraphClass.Caption), 26, true, "IncorrectCaptionFontSize"),
        new ForceJustificationLint(x => x.CaptionData is {Type: CaptionType.Figure}, [JustificationValues.Center], "FigureCaptionNotCentered"),
        // TODO: force no indent for captions?
        // TODO: empty line before figure and after caption?
        // TODO: force references to figures to be shortened.
        
        new ForceJustificationLint(IsOfClass(ConferenceParagraphClass.TableContent), [JustificationValues.Center], "TableContentNotCentered"),
        new ForceJustificationLint(x => x.CaptionData is {Type: CaptionType.Table}, [JustificationValues.Right, JustificationValues.End], "TableCaptionNotRightAligned"),
        
        new IncorrectCaptionedNumberingLint(_ => true, CaptionType.Figure, "IncorrectFigureNumbering", null, false),
        new IncorrectCaptionedNumberingLint(_ => true, CaptionType.Listing, "IncorrectListingNumbering", null, false),
        new IncorrectCaptionedNumberingLint(_ => true, CaptionType.Table, "IncorrectTableNumbering", null, false),
        
        new BibliographySourceNotReferencedLint(IsOfClass(ConferenceParagraphClass.BibliographySource))
    ];

    private static bool IsBodyText(ParagraphPropertiesTool tool)
        => tool.GetFeature(ConferenceParagraphData.Key)!.Class == ConferenceParagraphClass.BodyText;

    private static Predicate<ParagraphPropertiesTool> IsOfClass(ConferenceParagraphClass klass)
        => x => x.GetFeature(ConferenceParagraphData.Key)!.Class == klass;
}