using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Ntk;

public class NtkParagraphData
{
    public static readonly FeatureKey<NtkParagraphData, ParagraphPropertiesTool> Key = new();

    public ParagraphPropertiesTool Inner { get; }

    public NtkParagraphData(ParagraphPropertiesTool tool)
    {
        Inner = tool;
    }

    public bool IsUniversalDecimalClassifier { get; set; }
    public bool IsTitle { get; set; }
    public bool IsAuthorData { get; set; }
    public bool IsSourceInstitute { get; set; }
    public bool IsAbstract { get; set; }
    public bool IsKeywords { get; set; }
    public bool IsBibliographyHeader { get; set; }
    public bool IsBibliographySource { get; set; }
    
    public bool IsProbablyJunk { get; set; }

    public NtkParagraphClass Class
    {
        get
        {
            if (IsProbablyJunk) return NtkParagraphClass.Junk;
            if (IsUniversalDecimalClassifier) return NtkParagraphClass.UniversalDecimalClassifier;
            if (IsAuthorData) return NtkParagraphClass.AuthorDetails;
            if (IsSourceInstitute) return NtkParagraphClass.SourceInstitute;
            if (IsTitle) return NtkParagraphClass.ThesisTitle;
            if (IsAbstract) return NtkParagraphClass.Abstract;
            if (IsKeywords) return NtkParagraphClass.Keywords;
            if (IsBibliographyHeader) return NtkParagraphClass.BibliographyHeader;
            if (IsBibliographySource) return NtkParagraphClass.BibliographySource;
            if (Inner.CaptionData != null) return NtkParagraphClass.Caption;
            if (Inner.ProbablyCodeListing) return NtkParagraphClass.CodeListing;
            if (Inner.ContainingTableCell != null) return NtkParagraphClass.TableContent;

            return NtkParagraphClass.BodyText;
        }
    }

    public class Attacher : IClassifier
    {
        public void Classify(DocumentAnalysisContext ctx)
        {
            foreach (var p in ctx.AllParagraphs)
            {
                var tool = ctx.GetTool(p);
                tool.SetFeature(NtkParagraphData.Key, new NtkParagraphData(tool));
            }
        }
    }
}

public enum NtkParagraphClass
{
    UniversalDecimalClassifier,
    ThesisTitle,
    AuthorDetails,
    SourceInstitute,
    Abstract,
    Keywords,
    BodyText,
    Caption,
    BibliographyHeader,
    BibliographySource,
    CodeListing,
    TableContent,
    Junk,
}