using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class ParagraphIndentLint(Predicate<ParagraphPropertiesTool> predicate, int firstLine, int left, string firstLineId, string leftId) : ILint
{
    public void Run(LintContext ctx)
    {
        foreach (var p in ctx.Document.AllParagraphs)
        {
            if (string.IsNullOrWhiteSpace(Utils.CollectParagraphText(p, 10).Text))
            {
                continue;
            }

            ParagraphPropertiesTool tool = ctx.Document.GetTool(p);
            
            if (!predicate(tool)) continue; 
            
            if (tool.FirstLineIndent != firstLine)
            {
                ctx.AddMessage(new LintMessage(
                    firstLineId,
                    new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["ExpectedCm"] = Utils.TwipsToCm(firstLine).ToString(CultureInfo.CurrentCulture),
                            ["ActualCm"] = Utils.TwipsToCm(tool.FirstLineIndent ?? 0).ToString(CultureInfo.CurrentCulture),
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.Indentation == null)
                                p.ParagraphProperties.Indentation = new Indentation();

                            if (firstLine > 0)
                            {
                                p.ParagraphProperties.Indentation.FirstLine = firstLine.ToString();
                                p.ParagraphProperties.Indentation.Hanging = null;
                            } 
                            else if (firstLine < 0)
                            {
                                p.ParagraphProperties.Indentation.Hanging = (-firstLine).ToString();
                                p.ParagraphProperties.Indentation.FirstLine = null;
                            }
                            else
                            {
                                p.ParagraphProperties.Indentation.Hanging = null;
                                p.ParagraphProperties.Indentation.FirstLine = null;
                            }
                            
                        }
                    }
                );
            }
            
            if (!(tool.LeftIndent == left || (left == 0 && tool.LeftIndent == null)))
            {
                ctx.AddMessage(new LintMessage(
                        leftId,
                        new ParagraphDiagnosticContext(p))
                    {
                        Parameters = new()
                        {
                            ["ExpectedCm"] = Utils.TwipsToCm(left).ToString(CultureInfo.CurrentCulture),
                            ["ActualCm"] = Utils.TwipsToCm(tool.LeftIndent ?? 0).ToString(CultureInfo.CurrentCulture),
                        },
                        AutoFix = () =>
                        {
                            if (p.ParagraphProperties == null) p.ParagraphProperties = new ParagraphProperties();
                
                            if (ctx.GenerateRevisions) Utils.SnapshotParagraphProperties(p.ParagraphProperties);

                            if (p.ParagraphProperties.Indentation == null)
                                p.ParagraphProperties.Indentation = new Indentation();

                            p.ParagraphProperties.Indentation.Left = left.ToString();
                        }
                    }
                );
            }
        }
    }
}