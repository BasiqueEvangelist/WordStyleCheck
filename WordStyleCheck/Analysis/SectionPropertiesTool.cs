using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class SectionPropertiesTool
{
    private DocumentAnalysisContext _ctx;
    public SectionProperties Properties { get; }
    public List<Paragraph> Paragraphs { get; internal set; } = [];

    internal SectionPropertiesTool(DocumentAnalysisContext ctx, SectionProperties properties)
    {
        _ctx = ctx;
        Properties = properties;
    }

    public SectionMarkValues Type => Properties.GetFirstChild<SectionType>()?.Val?.Value ?? SectionMarkValues.NextPage;

    public Size? PageSize => Properties.GetFirstChild<PageSize>() is { Width.Value: var w, Height.Value: var h }
        ? new Size((int)w, (int)h)
        : null;

    public PageOrientationValues Orientation =>
        Properties.GetFirstChild<PageSize>()?.Orient?.Value ?? PageOrientationValues.Portrait;
}