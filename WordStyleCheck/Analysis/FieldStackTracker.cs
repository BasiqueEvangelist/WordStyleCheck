using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordStyleCheck.Analysis;

public static class FieldStackTracker
{
    public static Dictionary<OpenXmlElement, List<FieldStackEntry>> Run(Document doc)
    {
        Dictionary<OpenXmlElement, List<FieldStackEntry>> dict = [];

        if (doc.Body == null) return dict;

        List<FieldStackEntry> stack = [];

        Run(stack, dict, doc.Body);

        return dict;
    }

    private static void Run(List<FieldStackEntry> stack, Dictionary<OpenXmlElement, List<FieldStackEntry>> dict, OpenXmlElement element)
    {
        if (element is FieldChar fldChar)
        {
            if (fldChar.FieldCharType?.Value == FieldCharValues.Begin)
            {
                stack.Add(new FieldStackEntry());
            }
            else if (fldChar.FieldCharType?.Value == FieldCharValues.End)
            {
                stack.RemoveAt(stack.Count - 1);
            }
        } else if (element is FieldCode instrText)
        {
            stack[^1].InstrText = instrText.Text;
        }

        if (stack.Count > 0 && element is Paragraph)
        {
            dict.Add(element, [..stack]);
        }

        foreach (var child in element.ChildElements)
        {
            Run(stack, dict, child);
        }
        
        if (stack.Count > 0 && element is Paragraph)
        {
            var mapping = dict.GetValueOrDefault(element);

            if (mapping == null)
            {
                dict[element] = mapping = [];
            }

            foreach (var el in stack)
            {
                if (!mapping.Contains(el))
                    mapping.Add(el);
            }
        }
    }

    public class FieldStackEntry
    {
        public string? InstrText;
    }
}