using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public class Utils
{
    public static string ToPlainText(List<OpenXmlElement> elements)
    {
        // TODO: StringBuilder.
        string text = "";
        bool first = false;

        foreach (var el in elements)
        {
            if (!first) text += "\n";
            first = false;

            if (el is Paragraph p)
            {
                text += CollectParagraphText(p);
            }
        }

        return text;
    }
    
    public static string CollectParagraphText(Paragraph p) => CollectParagraphText(p, int.MaxValue).Text;
    
    public static (string Text, bool More) CollectParagraphText(Paragraph p, int needed)
    {
        StringBuilder neededText = new();
        bool more = false;
        
        bool CollectRunText(Run r) 
        {
            foreach (var text in r.ChildElements.OfType<Text>())
            {
                neededText.Append(text.Text);

                if (neededText.Length > needed)
                {
                    more = true;
                    return true;
                }
            }

            return false;
        }
        
        foreach (var c in p.ChildElements)
        {
            if (c is Run r)
            {
                if (CollectRunText(r))
                    goto outer;
            }
            else if (c is Hyperlink h)
            {
                foreach (var r2 in h.ChildElements.OfType<Run>())
                {
                    if (CollectRunText(r2))
                    {
                        goto outer;
                    }
                }
            }
        }
        outer:

        return (neededText.ToString().Substring(0, Math.Min(needed, neededText.Length)), more);
    }

    public static string CollectText(Run r)
    {
        StringBuilder text = new();

        foreach (var t in r.ChildElements.OfType<Text>())
        {
            text.Append(t.Text);
        }

        return text.ToString();
    }

    private static int _annotationIdCounter = 1;

    public static void StampTrackChange(TrackChangeType t)
    {
        t.Id = (_annotationIdCounter++).ToString();
        t.Author = "WordStyleCheck";
        t.Date = DateTimeValue.FromDateTime(DateTime.Now);
    }
    
    public static void StampTrackChange(RunTrackChangeType t)
    {
        t.Id = (_annotationIdCounter++).ToString();
        t.Author = "WordStyleCheck";
        t.Date = DateTimeValue.FromDateTime(DateTime.Now);
    }
    
    public static void SnapshotParagraphProperties(ParagraphProperties properties)
    {
        if (properties.ParagraphPropertiesChange?.Author == "WordStyleCheck") return;
        
        var old = (ParagraphProperties) properties.CloneNode(true);

        properties.ParagraphPropertiesChange = new ParagraphPropertiesChange
        {
            Author = "WordStyleCheck",
            Id = (_annotationIdCounter++).ToString(),
            Date = DateTimeValue.FromDateTime(DateTime.Now),
            ParagraphPropertiesExtended = new ParagraphPropertiesExtended(old.OuterXml)
        };
    }
    
    public static void SnapshotRunProperties(RunProperties properties)
    {
        if (properties.RunPropertiesChange?.Author == "WordStyleCheck") return;
        
        var old = (RunProperties) properties.CloneNode(true);

        properties.RunPropertiesChange = new RunPropertiesChange()
        {
            Author = "WordStyleCheck",
            Id = (_annotationIdCounter++).ToString(),
            Date = DateTimeValue.FromDateTime(DateTime.Now),
            PreviousRunProperties = new PreviousRunProperties(old.OuterXml)
        };
    }
    
    public static int? ParseTwipsMeasure(string? text)
    {
        if (text == null) return null;
        
        // TODO!!!: also work with measurement units. 
        return int.Parse(text);
    }
    
    public static int? ParseHpsMeasure(string? text)
    {
        if (text == null) return null;
        
        // TODO!!!: also work with measurement units. 
        return int.Parse(text);
    }

    public static T? AscendToAnscestor<T>(OpenXmlElement element) where T : OpenXmlCompositeElement
    {
        var parent = element.Parent;

        while (parent is not T and not null)
        {
            parent = parent.Parent;
        }

        return (T?)parent;
    }

    public static string Truncate(string text)
    {
        if (text.Length <= 20) return text;
        
        return text[..10] + "…" + text[^10..];
    }

    public static double TwipsToCm(double twips) => Math.Round(twips / 566.9291, 2);

    public static bool? ConvertOnOffType(OnOffType? value)
    {
        if (value == null) return null;
        
        return value.Val?.Value != false; 
    }
}
