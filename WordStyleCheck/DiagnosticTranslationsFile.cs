using System.Reflection;
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
                string dumped = key;

                if (parameters.Count > 0)
                {
                    dumped += " {";
                    dumped += string.Join(", ", parameters.Select(x => $"{x.Key} = '{x.Value}'"));
                    dumped += "}";
                }
                
                return [new Paragraph(new Run(new Text(dumped)))];
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

        public static DiagnosticTranslationsFile LoadEmbedded()
        {
            return LoadFromDocx(Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("WordStyleCheck.rules.docx")!);
        }
        
        public static DiagnosticTranslationsFile LoadFromDocx(Stream stream)
        {
            var translations = new Dictionary<string, List<OpenXmlElement>>();
            using (var doc = WordprocessingDocument.Open(stream, false))
            {
                string part = "";
                List<OpenXmlElement> current = [];
                foreach (var el in doc.MainDocumentPart!.Document!.Body!.ChildElements)
                {
                    if (el is Paragraph p)
                    {
                        CleanAndMergeRuns(p);
                        
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

        private static void CleanAndMergeRuns(Paragraph p)
        {
            var runs = p.ChildElements.OfType<Run>().ToList();
            
            foreach (var run in runs)
            {
                if (run.RunProperties is { } rPr)
                {
                    rPr.Languages = null;

                    if (rPr.ChildElements.Count == 0) run.RunProperties = null;
                }

                run.RsidRunAddition = null;
                run.RsidRunDeletion = null;
                run.RsidRunProperties = null;
            }

            for (int i = 1; i < runs.Count; i++)
            {
                if (runs[i - 1].RunProperties != null || runs[i].RunProperties != null) continue;

                var runChildren = runs[i].ChildElements.ToList();
                runs[i].RemoveAllChildren();
                runs[i - 1].Append(runChildren);
                runs[i].Remove();
                runs.RemoveAt(i);
                i -= 1;
            }

            foreach (var run in runs)
            {
                var texts = run.ChildElements.OfType<Text>().ToList();

                for (int i = 1; i < texts.Count(); i++)
                {
                    if (texts[i - 1].NextSibling() != texts[i]) continue;

                    texts[i - 1].Text += texts[i].Text;
                    texts[i - 1].Space = SpaceProcessingModeValues.Preserve;
                    texts[i].Remove();
                    texts.RemoveAt(i);
                    i -= 1;
                }
            }
        }
    }
}
