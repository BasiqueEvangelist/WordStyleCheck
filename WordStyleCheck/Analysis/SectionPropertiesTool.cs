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

    public PageMargins? PageMargins => Properties.GetFirstChild<PageMargin>() is
    {
        Top.Value: var top, Bottom.Value: var bottom, Left.Value: var left, Right.Value: var right,
        Header.Value: var header, Footer.Value: var footer, Gutter.Value: var gutter
    }
        ? new PageMargins(top, bottom, (int)left, (int)right, (int)header, (int)footer, (int)gutter)
        : null;
}

public readonly record struct PageMargins(int Top, int Bottom, int Left, int Right, int Header, int Footer, int Gutter)
{
    public bool CloseTo(PageMargins other)
    {
        return Math.Abs(Top - other.Top) < 5 && Math.Abs(Bottom - other.Bottom) < 5 &&
               Math.Abs(Left - other.Left) < 5 && Math.Abs(Right - other.Right) < 5 &&
               Math.Abs(Header - other.Header) < 5 && Math.Abs(Footer - other.Footer) < 5 &&
               Math.Abs(Gutter - other.Gutter) < 5;
    }
    
    public PageMargins Rotate()
    {
        return new(Left, Right, Bottom, Top, Header, Footer, Gutter);
    }
}