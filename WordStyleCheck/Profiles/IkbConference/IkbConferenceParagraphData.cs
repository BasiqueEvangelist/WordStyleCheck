using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.IkbConference;

public class IkbConferenceParagraphData
{
    public static readonly FeatureKey<IkbConferenceParagraphData, ParagraphPropertiesTool> Key = new();

    public ParagraphPropertiesTool Inner { get; }

    public IkbConferenceParagraphData(ParagraphPropertiesTool tool)
    {
        Inner = tool;
    }

    public List<string>? UniversalDecimalClassifier { get; set; }
    public AuthorData? AuthorData { get; set; }
    public bool IsTitle { get; set; }
    public bool IsAbstract { get; set; }
    public bool IsKeywords { get; set; }
    public bool IsBibliographyHeader { get; set; }
    public bool IsBibliographySource { get; set; }
    public bool IsCopyright { get; set; }

    public ConferenceParagraphClass Class
    {
        get
        {
            if (UniversalDecimalClassifier != null) return ConferenceParagraphClass.UniversalDecimalClassifier;
            if (AuthorData != null) return ConferenceParagraphClass.AuthorDetails;
            if (IsTitle) return ConferenceParagraphClass.ThesisTitle;
            if (IsAbstract) return ConferenceParagraphClass.Abstract;
            if (IsKeywords) return ConferenceParagraphClass.Keywords;
            if (IsBibliographyHeader) return ConferenceParagraphClass.BibliographyHeader;
            if (IsBibliographySource) return ConferenceParagraphClass.BibliographySource;
            if (IsCopyright) return ConferenceParagraphClass.Copyright;
            if (Inner.CaptionData != null) return ConferenceParagraphClass.Caption;
            if (Inner.ProbablyCodeListing) return ConferenceParagraphClass.CodeListing;
            if (Inner.ContainingTableCell != null) return ConferenceParagraphClass.TableContent;

            if (Inner.Justification == JustificationValues.Center && Inner.Runs.All(x => x.Bold))
                return ConferenceParagraphClass.Heading;

            return ConferenceParagraphClass.BodyText;
        }
    }

    public class Attacher : IClassifier
    {
        public void Classify(DocumentAnalysisContext ctx)
        {
            foreach (var p in ctx.AllParagraphs)
            {
                var tool = ctx.GetTool(p);
                tool.SetFeature(IkbConferenceParagraphData.Key, new IkbConferenceParagraphData(tool));
            }
        }
    }
}

public enum ConferenceParagraphClass
{
    UniversalDecimalClassifier,
    AuthorDetails,
    ThesisTitle,
    Abstract,
    Keywords,
    BodyText,
    Heading,
    Caption,
    BibliographyHeader,
    BibliographySource,
    Copyright,
    CodeListing,
    TableContent,
}