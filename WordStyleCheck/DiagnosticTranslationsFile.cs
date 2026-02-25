using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace WordStyleCheck
{
    public class DiagnosticTranslationsFile
    {
        private readonly Dictionary<string, List<OpenXmlElement>> _translations;

        private DiagnosticTranslationsFile(Dictionary<string, List<OpenXmlElement>> translations)
        {
            _translations = translations;
        }

        public List<OpenXmlElement> Translate(string key, Dictionary<string, string> parameters)
        {
            if (!_translations.TryGetValue(key, out var source))
            {
                return [new Paragraph(new Run(new Text(key)))];
            }

            var clonedList = source.Select(x => x.CloneNode(true)).ToList();
            
            // TODO: make this more efficient
            foreach (var el in clonedList)
            {
                foreach (var text in el.Descendants<Text>())
                {
                    foreach (var entry in parameters.OrderByDescending(x => x.Key))
                    {
                        text.Text = text.Text.Replace("$" + entry.Key, entry.Value);
                    }
                }
            }

            return clonedList;
        }
        
        public static DiagnosticTranslationsFile LoadFromDocx(string path)
        {
            var translations = new Dictionary<string, List<OpenXmlElement>>();
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                string part = "";
                List<OpenXmlElement> current = [];
                foreach (var el in doc.MainDocumentPart!.Document!.Body!.ChildElements)
                {
                    if (el is Paragraph p)
                    {
                        string text = Utils.CollectParagraphText(p);

                        if (text.StartsWith("@"))
                        {
                            if (text.StartsWith("@BEGIN "))
                            {
                                part = text.Substring("@BEGIN ".Length);
                                continue;
                            }
                            else if (text.StartsWith("@END"))
                            {
                                translations[part] = current;
                                current = [];
                                part = "";
                                continue;
                            }
                            else
                            {
                                throw new NotImplementedException("Incorrect structural command " + text);
                            }
                        }
                    }

                    if (part != "")
                    {
                        current.Add(el);
                    }
                }
            }

            return new DiagnosticTranslationsFile(translations);
        } 
    }
}
