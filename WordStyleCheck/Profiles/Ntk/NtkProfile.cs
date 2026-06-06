using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class NtkProfile : IProfile
{
    public string Id => "ntk";
    public string Name => "Научно-техническая конференция";
    public List<IClassifier> Classifiers { get; } = [
        new TableFigureClassifier(),
        new EquationTableClassifier(),
        new NtkParagraphData.Attacher(),
        new NtkPartsClassifier()
    ];
    public List<ILint> Lints { get; } = [
        new ForceContentsLint(
            IsOfClass(NtkParagraphClass.BibliographyHeader),
            _ => "Список источников и литературы",
            "IncorrectBibliographyHeaderContents"
        ),
        
        new ForbidLint(IsOfClass(NtkParagraphClass.Junk), "JunkParagraph"),
        
        new TextFontLint(),
        new FontSizeLint(
            IsOfClass(
                NtkParagraphClass.UniversalDecimalClassifier,
                NtkParagraphClass.ThesisTitle,
                NtkParagraphClass.AuthorDetails,
                NtkParagraphClass.SourceInstitute,
                NtkParagraphClass.BodyText,
                NtkParagraphClass.BibliographyHeader,
                NtkParagraphClass.BibliographySource
            ),
            28,
            true,
            "IncorrectMainFontSize"
        ),
        new ParagraphLineSpacingLint(_ => true, 288, "IncorrectLineSpacing"),
        new PageMarginsLint(new PageMargins(1134, 1134, 1134, 1134, 0, 0 ,0)),
        new PageSizeLint(),
        // TODO: footnotes.
        new ForceJustificationLint(IsOfClass(NtkParagraphClass.BodyText, NtkParagraphClass.BibliographyHeader, NtkParagraphClass.BibliographySource), [JustificationValues.Both], "MainTextNotJustified"),
        // automatic hyphenation
        // no orphan control?
        new ParagraphIndentLint(
            IsOfClass(
                NtkParagraphClass.Abstract,
                NtkParagraphClass.Keywords,
                NtkParagraphClass.BodyText,
                NtkParagraphClass.BibliographyHeader,
                NtkParagraphClass.BibliographySource
            ),
            709,
            0,
            "IncorrectFirstLineIndent",
            "IncorrectLeftIndent"
        ),
        // TODO: disallow listings.
        new IncorrectCaptionedNumberingLint(_ => true, CaptionType.Figure, "IncorrectFigureNumbering", "FigureNumberingMix"),
        // TODO: force formulae to be tables.
        // TODO: track numbering of formulae.
        // TODO: MathType support.
        // TODO: built-in Word math support.
        new ForceTableAutoWidthLint(),
        new IncorrectCaptionedNumberingLint(_ => true, CaptionType.Table, "IncorrectTableNumbering", "TableNumberingMix"),
        // TODO: (maybe) force tables to be at the bottom/top of the doc (if possible) 
        new WrongCaptionPositionLint(CaptionType.Table, false, "IncorrectTableCaptionPosition"),
        // TODO: force references to tables and figures to be abbreviated
        // TODO: (maybe) check bibliography?
        
        new ForceBoldLint(true, IsOfClass(NtkParagraphClass.UniversalDecimalClassifier), "UdcNotBold"),
        // TODO: force newlines between stuff.
        
        new ForceBoldLint(false, IsOfClass(NtkParagraphClass.BibliographyHeader), "BibliographyHeaderBold"),
        
        new IncorrectHeaderLint(IsOfClass(NtkParagraphClass.Abstract), ["Аннотация."], "Аннотация:", "IncorrectAbstractHeader"),
        new BoldItalicThenItalicLint(IsOfClass(NtkParagraphClass.Abstract), "Аннотация:", "AbstractHeaderMustBeBoldItalic", "AbstractBodyMustBeItalic"),
        new BoldItalicThenItalicLint(IsOfClass(NtkParagraphClass.Keywords), "Ключевые слова:", "KeywordsHeaderMustBeBoldItalic", "KeywordsBodyMustBeItalic"),
        new FontSizeLint(IsOfClass(NtkParagraphClass.Abstract, NtkParagraphClass.Keywords), 24, true, "IncorrectAbstractKeywordsFontSize"),
        
        new InstituteCapitalizationLint(),
        
        new TextColorLint(),
        new BadOuterWhitespaceLint(_ => true),
        
        new QuoteTrackerLint(IsOfClass(NtkParagraphClass.BodyText))
    ];
    
    private static Predicate<ParagraphPropertiesTool> IsOfClass(params NtkParagraphClass[] klass)
        => x => klass.Contains(x.GetFeature(NtkParagraphData.Key)!.Class);
}