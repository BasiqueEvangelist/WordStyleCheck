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
        new TextColorLint(),
        
        new ForceContentsLint(
            IsOfClass(NtkParagraphClass.BibliographyHeader),
            _ => "Список источников и литературы",
            "IncorrectBibliographyHeaderContents"
        ),
        
        new ForceContentsLint(
            IsOfClass(NtkParagraphClass.TableCaption, NtkParagraphClass.FigureCaption),
            x =>
            {
                var data = x.CaptionData!.Value;
                
                string desc = data.GetDesc(x.Contents);

                string correct = (data.Type, data.IsContinuation) switch
                {
                    (CaptionType.Figure, _) => "Рисунок ",
                    (CaptionType.Listing, _) => "Листинг ",
                    (CaptionType.Table, false) => "Таблица ",
                    (CaptionType.Table, true) => "Продолжение табл. ",
                    _ => throw new ArgumentOutOfRangeException()
                };

                correct += data.Number;
                correct += ".";

                if (desc != "")
                {
                    correct += " " + desc;
                }

                return correct;
            },
            "IncorrectCaptionText"
        ),
        
        new NoUdcLint(),
        
        new ForbidLint(IsOfClass(NtkParagraphClass.Junk), "JunkParagraph"),
        
        new ZapFootnotesEndnotesLint(IsOfClass(NtkParagraphClass.BibliographySource)),
        
        new TextFontLint(),
        new FontSizeLint(
            IsOfClass(
                NtkParagraphClass.UniversalDecimalClassifier,
                NtkParagraphClass.ThesisTitle,
                NtkParagraphClass.AuthorDetails,
                NtkParagraphClass.SupervisorDetails,
                NtkParagraphClass.ConsultantDetails,
                NtkParagraphClass.SourceInstitute,
                NtkParagraphClass.Heading,
                NtkParagraphClass.BodyText,
                NtkParagraphClass.BibliographyHeader,
                NtkParagraphClass.BibliographySource
            ),
            28,
            true,
            "IncorrectMainFontSize"
        ),
        new ParagraphLineSpacingLint(
            IsOfClass(
                NtkParagraphClass.UniversalDecimalClassifier,
                NtkParagraphClass.SourceInstitute,
                NtkParagraphClass.Heading,
                NtkParagraphClass.BodyText,
                NtkParagraphClass.FigureCaption,
                NtkParagraphClass.TableCaption,
                NtkParagraphClass.BibliographyHeader,
                NtkParagraphClass.BibliographySource,
                NtkParagraphClass.CodeListing,
                NtkParagraphClass.TableContent
            ),
            288,
            "IncorrectLineSpacing"
        ),
        new ParagraphLineSpacingLint(
            IsOfClass(
                NtkParagraphClass.ThesisTitle,
                NtkParagraphClass.AuthorDetails,
                NtkParagraphClass.SupervisorDetails,
                NtkParagraphClass.ConsultantDetails,
                NtkParagraphClass.Abstract,
                NtkParagraphClass.Keywords
            ),
            276,
            "IncorrectLineSpacing"
        ),
        new PageMarginsLint(new PageMargins(1134, 1134, 1134, 1134, 0, 0 ,0)),
        new PageSizeLint(),
        // TODO: footnotes.
        new ForceJustificationLint(IsOfClass(NtkParagraphClass.BodyText, NtkParagraphClass.Heading, NtkParagraphClass.BibliographyHeader, NtkParagraphClass.BibliographySource), [JustificationValues.Both], "MainTextNotJustified"),
        // automatic hyphenation
        // no orphan control?
        new ParagraphIndentLint(
            IsOfClass(
                NtkParagraphClass.Abstract,
                NtkParagraphClass.Keywords,
                NtkParagraphClass.Heading,
                NtkParagraphClass.BodyText,
                NtkParagraphClass.BibliographyHeader,
                NtkParagraphClass.BibliographySource
            ),
            709,
            0, 0,
            "IncorrectFirstLineIndent",
            "IncorrectLeftIndent",
            "IncorrectRightIndent"),
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
        new ParagraphIndentLint(IsOfClass(NtkParagraphClass.UniversalDecimalClassifier), 0, 0, 0, "IncorrectUdcFirstLineIndent", "IncorrectUdcLeftIndent", "IncorrectUdcRightIndent"),
        // TODO: force newlines between stuff.
        
        new ForceJustificationLint(IsOfClass(NtkParagraphClass.ThesisTitle, NtkParagraphClass.AuthorDetails, NtkParagraphClass.SupervisorDetails, NtkParagraphClass.ConsultantDetails, NtkParagraphClass.SourceInstitute), [JustificationValues.Center], "HeaderNotCentered"),
        new ParagraphIndentLint(IsOfClass(NtkParagraphClass.ThesisTitle, NtkParagraphClass.AuthorDetails, NtkParagraphClass.SupervisorDetails, NtkParagraphClass.ConsultantDetails, NtkParagraphClass.SourceInstitute), 0, 0, 0, "IncorrectHeaderFirstLineIndent", "IncorrectHeaderLeftIndent", "IncorrectHeaderRightIndent"),
        new ForceCapsLint(IsOfClass(NtkParagraphClass.ThesisTitle), "TitleMustBeCaps"),
        
        new SupervisorFixerLint(),
        new ForceBoldLint(true, IsOfClass(NtkParagraphClass.SupervisorDetails), "SupervisorNotBold"),
        
        new ForceBoldLint(false, IsOfClass(NtkParagraphClass.BibliographyHeader), "BibliographyHeaderBold"),

        new ForceBoldLint(false, IsOfClass(NtkParagraphClass.Heading), "HeadingBold"),
        new ForceItalicLint(true, IsOfClass(NtkParagraphClass.Heading), "HeadingNotItalic"),
        
        new ForceBoldLint(true, IsOfClass(NtkParagraphClass.SourceInstitute), "SourceInstituteNotBold"),
        new ForceItalicLint(true, IsOfClass(NtkParagraphClass.SourceInstitute), "SourceInstituteNotItalic"),
        
        new IncorrectHeaderLint(IsOfClass(NtkParagraphClass.Abstract), ["Аннотация. ", "Аннотация.", "Аннотация:"], "Аннотация: ", "IncorrectAbstractHeader"),
        new BoldItalicThenItalicLint(IsOfClass(NtkParagraphClass.Abstract), "Аннотация:", "AbstractHeaderMustBeBoldItalic", "AbstractBodyMustBeItalic"),
        new BoldItalicThenItalicLint(IsOfClass(NtkParagraphClass.Keywords), "Ключевые слова:", "KeywordsHeaderMustBeBoldItalic", "KeywordsBodyMustBeItalic"),
        new FontSizeLint(IsOfClass(NtkParagraphClass.Abstract, NtkParagraphClass.Keywords), 24, true, "IncorrectAbstractKeywordsFontSize"),
        new AbstractStartsWithLowercaseLint(),
        
        new ForceJustificationLint(IsOfClass(NtkParagraphClass.FigureCaption), [JustificationValues.Center], "FigureCaptionMustBeCentered"),
        new ForceJustificationLint(IsOfClass(NtkParagraphClass.TableCaption), [JustificationValues.Right], "TableCaptionMustBeRightAligned"),
        new FontSizeLint(IsOfClass(NtkParagraphClass.FigureCaption, NtkParagraphClass.TableCaption), 26, true, "IncorrectCaptionFontSize"),
        new ParagraphIndentLint(IsOfClass(NtkParagraphClass.FigureCaption), 0, 0, 0, "IncorrectFigureCaptionIndent", "IncorrectFigureCaptionIndent", "IncorrectFigureCaptionIndent"),
        
        new InstituteLint(),
        
        new StripAuthorJunkLint(),
        new ForceBoldLint(true, IsOfClass(NtkParagraphClass.AuthorDetails), "AuthorNotBold"),
        
        new BadOuterWhitespaceLint(_ => true),
        
        new QuoteTrackerLint(IsOfClass(NtkParagraphClass.BodyText)),
        
        new TableifyFiguresLint(),
        
        new EmptyLineControlLint(
            [
                new EmptyLineControlLint.Rule(
                    IsOfClass(NtkParagraphClass.Abstract),
                    "NoEmptyBeforeAbstract",
                    28
                ),
                new EmptyLineControlLint.Rule(
                    IsOfClass(NtkParagraphClass.BibliographyHeader),
                    "NoEmptyBeforeBibliography",
                    28
                )
            ],
            [
                new EmptyLineControlLint.Rule(
                    IsOfClass(NtkParagraphClass.Keywords),
                    "NoEmptyAfterKeywords",
                    24
                ),
                new EmptyLineControlLint.Rule(
                    x => x is TablePropertiesTool {Class: TableClass.Table or TableClass.TableContinuation or TableClass.Figure},
                    "NoEmptyAfterTable",
                    28
                )
            ],
            "ForbiddenEmpty"
        ),
        
        new NoInterParagraphSpacingLint(),
        
        new ForbidPageBreaksLint(),
    ];
    
    private static Predicate<IBlockLevelPropertiesTool> IsOfClass(params NtkParagraphClass[] klass)
        => x => x is ParagraphPropertiesTool p && klass.Contains(p.GetFeature(NtkParagraphData.Key)!.Class);
}