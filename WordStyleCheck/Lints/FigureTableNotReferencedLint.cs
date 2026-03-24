using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;
using WordStyleCheck.Context;

namespace WordStyleCheck.Lints;

public class FigureTableNotReferencedLint : ILint
{
    public IReadOnlyList<string> EmittedDiagnostics { get; } =
    [
        "FigureBeforeFirstReference", "FigureNotReferenced",
        "TableBeforeFirstReference", "TableNotReferenced",
        "ListingBeforeFirstReference", "ListingNotReferenced"
    ];
    
    public void Run(LintContext ctx)
    {
        Dictionary<string, Paragraph> referencedFigureNumbers = [];
        Dictionary<string, Paragraph> referencedTableNumbers = [];
        Dictionary<string, Paragraph> referencedListingNumbers = [];

        foreach (var other in ctx.Document.AllParagraphs)
        {
            var oTool = ctx.Document.GetTool(other);
            
            if (oTool.Class == ParagraphClass.Caption) continue;

            var text = oTool.Contents;

            int i = 0;

            string ConsumeWord()
            {
                int wStart = i;
                while (i < text.Length && (char.IsLetter(text[i]) || text[i] == '.'))
                {
                    i += 1;
                }

                return text.Substring(wStart, i - wStart);
            }
            
            string? ConsumeObjectNumber()
            {
                int nStart = i;

                if (i >= text.Length) return null;

                if (char.IsLetter(text[i]))
                {
                    i++;
                    if (i >= text.Length || text[i] != '.') return null;
                    i++;
                }
                
                if (i >= text.Length || !char.IsDigit(text[i])) return null;
                
                while (i < text.Length && (char.IsDigit(text[i]) || text[i] == '.')) i++;


                return text.Substring(nStart, i - nStart).TrimEnd('.');
            }

            void ConsumeWhitespace()
            {
                while (i < text.Length && char.IsWhiteSpace(text[i])) i += 1;
            }

            void AddRange(string start, string end, Dictionary<string, Paragraph> references)
            {
                string[] startSplit = start.Split(".");
                string[] endSplit = end.Split(".");
                        
                if (startSplit.Length != endSplit.Length) return;
                if (!startSplit[..^1].SequenceEqual(endSplit[..^1])) return;
                        
                if (!int.TryParse(startSplit[^1], out var startNum)) return;
                if (!int.TryParse(endSplit[^1], out var endNum)) return;
                        
                for (int j = startNum; j <= endNum; j++)
                {
                    startSplit[^1] = j.ToString();
                        
                    string referenced = string.Join(".", startSplit);
                        
                    references.TryAdd(referenced, other);
                }
            }

            void ConsumeReferences(Dictionary<string, Paragraph> references)
            {
                ConsumeWhitespace();

                while (true)
                {
                    string? start = ConsumeObjectNumber();
                    
                    if (start == null) return;
                    
                    references.TryAdd(start, other);
                    
                    ConsumeWhitespace();
                    
                    if (i < text.Length && text[i] is '-' or '–')
                    {
                        i += 1;
                        ConsumeWhitespace();
                        string? end = ConsumeObjectNumber();
                        if (end != null)
                        {
                            AddRange(start, end, references);
                        }

                        ConsumeWhitespace();
                    }

                    if (i < text.Length && text[i] == ',') i += 1;
                    
                    ConsumeWhitespace();
                }
            }
            
            while (i < text.Length)
            {
                string word = ConsumeWord();

                if (word.ToLower() is "рис." or "рисунок" or "рисунки" or "рисунке" or "рисунком" or "рисунках")
                {
                    ConsumeReferences(referencedFigureNumbers);
                } 
                else if (word.ToLower() is "табл." or "таблица" or "таблицы" or "таблице" or "таблицу")
                {
                    ConsumeReferences(referencedTableNumbers);
                }
                else if (word.ToLower() is "листинг" or "листинги" or "листенге" or "листингах")
                {
                    ConsumeReferences(referencedListingNumbers);
                }

                if (string.IsNullOrWhiteSpace(word)) i += 1;
            }
        }

        foreach (var p in ctx.Document.AllParagraphs)
        {
            var tool = ctx.Document.GetTool(p);

            void EmitDiagnosticsIfNeeded(Dictionary<string, Paragraph> referencedNumbers, string beforeFirstRefId,
                string notRefId)
            {
                if (referencedNumbers.TryGetValue(tool.CaptionData!.Value.Number, out var firstMention))
                {
                    var fToplevel = FindTopLevel(p);
                    var mToplevel = FindTopLevel(firstMention);

                    var bodyList = fToplevel.Parent!.ChildElements.ToList();

                    var fIdx = bodyList.IndexOf(fToplevel);
                    var mIdx = bodyList.IndexOf(mToplevel);

                    if (fIdx < mIdx)
                    {
                        ctx.AddMessage(new LintMessage(beforeFirstRefId, new ParagraphDiagnosticContext(p)));
                    }
                }
                else
                {
                    ctx.AddMessage(new LintMessage(notRefId, new ParagraphDiagnosticContext(p)));
                }
            }
            
            if ((tool.CaptionData?.Type) == CaptionType.Figure)
            {
                EmitDiagnosticsIfNeeded(referencedFigureNumbers, "FigureBeforeFirstReference", "FigureNotReferenced");
            }
            else if (tool is { CaptionData: { Type: CaptionType.Table, IsContinuation: false } })
            {
                EmitDiagnosticsIfNeeded(referencedTableNumbers, "TableBeforeFirstReference", "TableNotReferenced");
            }
            else if (tool is { CaptionData.Type: CaptionType.Listing })
            {
                EmitDiagnosticsIfNeeded(referencedListingNumbers, "ListingBeforeFirstReference", "ListingNotReferenced");
            }
        }
    }

    private static OpenXmlElement FindTopLevel(OpenXmlElement e)
    {
        while (e.Parent is not (null or Body)) e = e.Parent;

        return e;
    }
}