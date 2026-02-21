using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck;

public class Utils
{
    public static string CollectParagraphText(Paragraph p) => CollectParagraphText(p, int.MaxValue).Text;
    
    public static (string Text, bool More) CollectParagraphText(Paragraph p, int needed)
    {
        StringBuilder neededText = new();
        bool more = false;
        foreach (var text in p.Descendants<Text>())
        {
            neededText.Append(text.Text);

            if (neededText.Length > needed)
            {
                more = true;
                break;
            }
        }

        return (neededText.ToString(), more);
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
    
    public static int? ParseTwipsMeasure(string? text)
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
}
