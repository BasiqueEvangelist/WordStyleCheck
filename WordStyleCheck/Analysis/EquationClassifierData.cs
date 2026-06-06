using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public class EquationClassifierData
{
    public required OpenXmlElement MathElement { get; init; }
    public required string? Number { get; init; }
    
    public static EquationClassifierData? Classify(ParagraphPropertiesTool p)
    {
        OpenXmlElement? oMath = null;
        
        foreach (var c1 in p.Paragraph.ChildElements)
        {
            if (c1 is Run r)
            {
                foreach (var c2 in r.ChildElements)
                {
                    if (c2 is Text t && !string.IsNullOrWhiteSpace(t.Text))
                        return null;
                }
            }

            if (c1 is DocumentFormat.OpenXml.Math.OfficeMath or DocumentFormat.OpenXml.Math.Paragraph)
            {
                oMath = c1;
                break;
            }

            if (oMath != null) break;
        }

        if (oMath == null) return null;

        string contents = p.Contents.Trim();

        string? number;

        if (string.IsNullOrWhiteSpace(contents))
        {
            number = null;
        } else if (contents.Length > 0 && contents[0] == '(' && contents[^1] == ')')
        {
            number = contents.Substring(1, contents.Length - 2);
        }
        else
        {
            return null;
        }

        return new EquationClassifierData
        {
            MathElement = oMath,
            Number = number
        };
    }
}