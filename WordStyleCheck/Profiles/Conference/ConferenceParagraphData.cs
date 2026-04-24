using WordStyleCheck.Analysis;

namespace WordStyleCheck.Profiles.Conference;

public class ConferenceParagraphData
{
    public static readonly FeatureKey<ConferenceParagraphData, ParagraphPropertiesTool> Key = new();

    public ParagraphPropertiesTool Inner { get; }

    public ConferenceParagraphData(ParagraphPropertiesTool tool)
    {
        Inner = tool;
    }

    public ConferenceParagraphClass Class { get; set; } = ConferenceParagraphClass.BodyText;
}

public enum ConferenceParagraphClass
{
    UniversalDecimalClassifier,
    AuthorName,
    AuthorDetails,
    AuthorEmail,
    ThesisName,
    Abstract,
    Keywords,
    BodyText,
    Caption,
    BibliographyHeader,
    BibliographySource,
    Copyright,
}