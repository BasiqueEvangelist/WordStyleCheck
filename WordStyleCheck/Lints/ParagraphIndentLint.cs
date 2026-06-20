using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphIndentLint(
    Predicate<ParagraphPropertiesTool> predicate,
    int firstLine,
    int left,
    int right,
    string firstLineId,
    string leftId,
    string rightId) : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } = [firstLineId, leftId];

    public void Run(ILintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;
            if (!predicate(tool)) continue; 
            
            if (Math.Abs((tool.FirstLineIndent ?? 0) - firstLine) >= 5 && !(tool.MaybeParagraphContinuation && (tool.FirstLineIndent ?? 0) == 0))
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(
                            firstLineId,
                            DiagnosticType.FormattingError,
                            new ParagraphDiagnosticContext(p))
                        {
                            Parameters = new()
                            {
                                ["ExpectedCm"] = Utils.TwipsToCm(firstLine).ToString(CultureInfo.InvariantCulture),
                                ["ActualCm"] = Utils.TwipsToCm(tool.FirstLineIndent ?? 0)
                                    .ToString(CultureInfo.InvariantCulture),
                            }
                        }
                    );
                }
                else
                {
                    ctx.MarkAutoFixed();
                    
                    if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    if (p.ParagraphProperties.Indentation == null)
                        p.ParagraphProperties.Indentation = new Indentation();

                    if (firstLine > 0)
                    {
                        p.ParagraphProperties.Indentation.FirstLine = firstLine.ToString();
                        p.ParagraphProperties.Indentation.Hanging = null;
                        p.ParagraphProperties.Indentation.FirstLineChars = null;
                    } 
                    else if (firstLine < 0)
                    {
                        p.ParagraphProperties.Indentation.Hanging = (-firstLine).ToString();
                        p.ParagraphProperties.Indentation.FirstLine = null;
                        p.ParagraphProperties.Indentation.FirstLineChars = null;
                    }
                    else
                    {
                        p.ParagraphProperties.Indentation.FirstLine = "0";
                        p.ParagraphProperties.Indentation.Hanging = null;
                        p.ParagraphProperties.Indentation.FirstLineChars = null;
                    }
                }
            }
        }

        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);

            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;
            if (!predicate(tool)) continue;
            
            if (Math.Abs((tool.LeftIndent ?? 0) - left) >= 5)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(
                            leftId,
                            DiagnosticType.FormattingError,
                            new ParagraphDiagnosticContext(p))
                        {
                            Parameters = new()
                            {
                                ["ExpectedCm"] = Utils.TwipsToCm(left).ToString(CultureInfo.InvariantCulture),
                                ["ActualCm"] = Utils.TwipsToCm(tool.LeftIndent ?? 0)
                                    .ToString(CultureInfo.InvariantCulture),
                            }
                        }
                    );
                }
                else
                {
                    ctx.MarkAutoFixed();

                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.Indentation ??= new Indentation();
                    p.ParagraphProperties.Indentation.Left = left.ToString();
                }
            }
        }
        
        foreach (var p in ctx.Document.AllParagraphs)
        {
            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);

            if (tool.IsEmptyOrDrawing) continue;
            if (tool.IsIgnored) continue;
            if (!predicate(tool)) continue;
            
            if (Math.Abs((tool.RightIndent ?? 0) - right) >= 5)
            {
                if (!ctx.AutomaticallyFix)
                {
                    ctx.AddMessage(new LintDiagnostic(
                            rightId,
                            DiagnosticType.FormattingError,
                            new ParagraphDiagnosticContext(p))
                        {
                            Parameters = new()
                            {
                                ["ExpectedCm"] = Utils.TwipsToCm(right).ToString(CultureInfo.InvariantCulture),
                                ["ActualCm"] = Utils.TwipsToCm(tool.RightIndent ?? 0)
                                    .ToString(CultureInfo.InvariantCulture),
                            }
                        }
                    );
                }
                else
                {
                    ctx.MarkAutoFixed();

                    p.ParagraphProperties ??= new ParagraphProperties();
                    if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                    p.ParagraphProperties.Indentation ??= new Indentation();
                    p.ParagraphProperties.Indentation.Right = right.ToString();
                }
            }
        }
    }
}