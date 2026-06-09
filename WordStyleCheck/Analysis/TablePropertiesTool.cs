using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class TablePropertiesTool : IBlockLevelPropertiesTool
{
    private DocumentAnalysisContext _ctx;
    public Table Table { get; }
    OpenXmlElement IBlockLevelPropertiesTool.Element => Table;
    public List<TableRowTool> Rows { get; }
    public int ColumnCount { get; }

    public ParagraphPropertiesTool? Caption { get; internal set; }
    public EquationClassifierData? EquationData { get; set; }
    public Style? Style { get; }
    public bool IsOutsideOfDocument { get; set; }
    
    public TableRowTool this[int index] => Rows[index];

    internal TablePropertiesTool(DocumentAnalysisContext ctx, Table table)
    {
        _ctx = ctx;
        Table = table;

        Rows = Table.ChildElements.OfType<TableRow>().Select(x => new TableRowTool(x)).ToList();
        ColumnCount = Rows.Count > 0 ? Rows[0].Cells.Count : 0;
        
        string? styleId = table.TableProperties?.TableStyle?.Val?.Value;
        
        Style = new Func<Style?>(() =>
        {
            if (styleId != null)
            {
                return ctx.GetStyle(StyleValues.Table, styleId);
            }

            return null;
        })();
    }

    public TableWidthUnitValues? WidthType => FollowPropertyChain(
        x => x.TableWidth?.Type?.Value,
        x => null
    ) ?? TableWidthUnitValues.Auto;

    public TableClass Class
    {
        get
        {
            if (IsOutsideOfDocument) return TableClass.ManualLayout;
            if (Rows.Count == 0 || ColumnCount == 0) return TableClass.ManualLayout;
            
            if (Caption is {CaptionData: {} data})
            {
                return data.Type switch
                {
                    CaptionType.Table => data.IsContinuation ? TableClass.TableContinuation : TableClass.Table,
                    CaptionType.Figure => TableClass.Figure,
                    CaptionType.Listing => TableClass.Listing,
                    _ => throw new NotImplementedException()
                };
            }

            if (EquationData != null) return TableClass.DisplayEquation;
            
            return TableClass.Unknown;
        }
    }
    
    private T? FollowPropertyChain<T>(Func<TableProperties, T?> getter, Func<StyleTableProperties, T?> styleGetter)
    {
        if (Table.TableProperties != null)
        {
            var result = getter(Table.TableProperties);
            if (result != null)
                return result;
        }
        
        if (Style?.StyleId != null)
        {
            var result = _ctx.FollowStyleChain(StyleValues.Table, Style.StyleId, x => x.StyleTableProperties != null ? styleGetter(x.StyleTableProperties) : default);
            if (result != null)
                return result;
        }

        return default;
    }
}

public class TableRowTool
{
    public TableRow Row { get; }
    public List<TableCell> Cells { get; }
    
    internal TableRowTool(TableRow row)
    {
        Row = row;
        Cells = Row.ChildElements.OfType<TableCell>().ToList();
    }

    public TableCell this[int index] => Cells[index];
}

public enum TableClass
{
    Unknown,
    Table,
    TableContinuation,
    DisplayEquation,
    Listing,
    Figure,
    ManualLayout
}