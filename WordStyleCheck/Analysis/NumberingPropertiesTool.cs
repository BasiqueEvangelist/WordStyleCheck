using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class NumberingPropertiesTool : INumbering
{
    private DocumentAnalysisContext _ctx;
    public NumberingInstance Numbering { get; }

    internal NumberingPropertiesTool(DocumentAnalysisContext ctx, NumberingInstance numbering)
    {
        _ctx = ctx;
        Numbering = numbering;
    }

    public List<Paragraph> Paragraphs { get; } = [];
}