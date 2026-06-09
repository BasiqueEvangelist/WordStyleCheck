using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Ntk;

public class TableifyFiguresLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = ["FigureMustBeTable"];
    
    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);
            
            if (tool.CaptionData is not {Type: CaptionType.Figure, TargetedElement: Paragraph targeted}) continue;

            if (!ctx.AutomaticallyFix)
            {
                ctx.AddMessage(new LintDiagnostic("FigureMustBeTable", DiagnosticType.FormattingError, new ParagraphDiagnosticContext([targeted, p])));
            }
            else
            {
                ctx.MarkAutoFixed();
                
                Table t = new()
                {
                    TableProperties = new()
                    {
                        TableWidth = new()
                        {
                            Type = TableWidthUnitValues.Auto,
                            Width = "0"
                        },
                        TableJustification = new()
                        {
                            Val = TableRowAlignmentValues.Center
                        },
                        TableBorders = new()
                        {
                            TopBorder = new()
                            {
                                Val = BorderValues.None
                            },
                            LeftBorder = new()
                            {
                                Val = BorderValues.None
                            },
                            BottomBorder = new()
                            {
                                Val = BorderValues.None
                            },
                            RightBorder = new()
                            {
                                Val = BorderValues.None
                            },
                            InsideHorizontalBorder = new()
                            {
                                Val = BorderValues.None
                            },
                            InsideVerticalBorder = new()
                            {
                                Val = BorderValues.None
                            }
                        }
                    },
                    TableGrid = new TableGrid()
                };

                targeted.InsertBeforeSelf(t);
                
                targeted.Remove();
                p.Remove();

                t.Append(
                    new TableRow(
                        new TableRowProperties(
                            new TableJustification() {Val = TableRowAlignmentValues.Center}
                        ),
                        new TableCell(
                            targeted
                        )
                    ),
                    new TableRow(
                        new TableRowProperties(
                            new TableJustification() {Val = TableRowAlignmentValues.Center}
                        ),
                        new TableCell(
                            p
                        )
                    )
                );

                t.InsertAfterSelf(new Paragraph(new Run()));
                
                targeted.ParagraphProperties ??= new ParagraphProperties();
                if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(targeted.ParagraphProperties);

                targeted.ParagraphProperties.Justification ??= new Justification();
                targeted.ParagraphProperties.Justification.Val = JustificationValues.Center;

                foreach (var anchor in targeted.Descendants<Anchor>().ToList())
                {
                    var inline = new Inline();

                    foreach (var child in anchor.ChildElements.ToList())
                    {
                        child.Remove();
                        inline.AppendChild(child);
                    }

                    anchor.InsertBeforeSelf(inline);
                    anchor.Remove();
                }

                ctx.Document.GetTool(t).Caption = tool;
            }
        }
    }
}