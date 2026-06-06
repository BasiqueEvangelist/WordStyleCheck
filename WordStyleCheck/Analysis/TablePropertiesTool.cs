using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class TablePropertiesTool
{
    private DocumentAnalysisContext _ctx;
    public Table Table { get; }
    public List<TableRowTool> Rows { get; }
    public int ColumnCount { get; }

    public ParagraphPropertiesTool? Caption { get; internal set; }
    public EquationClassifierData? EquationData { get; set; }
    public bool IsOutsideOfDocument { get; set; }
    
    public TableRowTool this[int index] => Rows[index];

    internal TablePropertiesTool(DocumentAnalysisContext ctx, Table table)
    {
        _ctx = ctx;
        Table = table;

        Rows = Table.ChildElements.OfType<TableRow>().Select(x => new TableRowTool(x)).ToList();
        ColumnCount = Rows.Count > 0 ? Rows[0].Cells.Count : 0;
    }

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