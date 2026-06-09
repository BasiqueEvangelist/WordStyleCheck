using DocumentFormat.OpenXml;

namespace WordStyleCheck.Analysis;

public interface IBlockLevelPropertiesTool
{
    public OpenXmlElement Element { get; }
}