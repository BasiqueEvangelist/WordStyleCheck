using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public interface INumbering
{
    List<Paragraph> Paragraphs { get; }
}