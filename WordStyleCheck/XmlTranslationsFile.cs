using System.Reflection;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordStyleCheck.Analysis;

namespace WordStyleCheck;

public class XmlTranslationsFile
{
    private readonly Dictionary<string, XElement> _translations;

    private XmlTranslationsFile(Dictionary<string, XElement> translations)
    {
        _translations = translations;
    }

    public static XmlTranslationsFile LoadEmbedded()
    {
        XDocument doc = XDocument.Load(Assembly.GetExecutingAssembly()
            .GetManifestResourceStream("WordStyleCheck.rules.xml")!, LoadOptions.None);
        return LoadFrom(doc);
    }

    public static XmlTranslationsFile LoadFrom(XDocument doc)
    {
        if (doc.Root!.Name != "dict")
            throw new InvalidDataException("Root element of translations file must be <dict>");

        Dictionary<string, XElement> translations = [];
        
        foreach (var el in doc.Root.Elements())
        {
            if (el.Name != "str")
                throw new InvalidDataException("Child element of <dict> must be <str>");

            string? key = el.Attribute("key")?.Value;
            if (key == null || string.IsNullOrWhiteSpace(key))
                throw new InvalidDataException("<str> must have a valid key");

            translations[key] = el;
        }

        return new XmlTranslationsFile(translations);
    }

    private string _processParameters(string text, Dictionary<string, string> parameters)
    {
        foreach (var entry in parameters.OrderByDescending(x => x.Key))
        {
            text = text.Replace("${" + entry.Key + "}", entry.Value);
        }

        return text;
    }

    public List<OpenXmlElement> Translate(string key, Dictionary<string, string> parameters, DocumentAnalysisContext? doc)
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

        Text currentText = new Text("");
        Run currentRun = new Run(currentText);
        Hyperlink? currentLink = null;
        Paragraph currentParagraph = new Paragraph(currentRun);
        List<OpenXmlElement> paragraphs = [currentParagraph];

        void AddRun(Run r)
        {
            if (currentLink != null)
            {
                currentLink.Append(r);
            }
            else
            {
                currentParagraph.Append(r);
            }
        }

        void ProcessChildrenOf(XElement el, bool bold, bool italic, Uri? url)
        {
            var nodeList = el.Nodes().ToList();
            
            for (var i = 0; i < nodeList.Count; i++)
            {
                var child = nodeList[i];
                
                if (child is XText text)
                {
                    string t = text.Value;

                    if (i == 0) t = t.TrimStart();
                    if (i == nodeList.Count - 1) t = t.TrimEnd();
                    
                    currentText.Text += _processParameters(t, parameters);
                }
                else if (child is XElement subEl)
                {
                    bool needReset = false;
                    if (subEl.Name == "b")
                    {
                        if (!bold)
                        {
                            needReset = true;

                            currentText = new Text("");
                            currentRun = new Run(new RunProperties(), currentText);

                            currentRun.RunProperties!.Bold = new Bold();
                            if (italic) currentRun.RunProperties!.Italic = new Italic();

                            AddRun(currentRun);
                        }

                        ProcessChildrenOf(subEl, true, italic, url);

                        if (needReset)
                        {
                            currentText = new Text("");
                            currentRun = new Run(new RunProperties(), currentText);

                            if (bold) currentRun.RunProperties!.Bold = new Bold();
                            if (italic) currentRun.RunProperties!.Italic = new Italic();

                            AddRun(currentRun);
                        }
                    }
                    else if (subEl.Name == "i")
                    {
                        if (!italic)
                        {
                            needReset = true;

                            currentText = new Text("");
                            currentRun = new Run(new RunProperties(), currentText);

                            if (bold) currentRun.RunProperties!.Bold = new Bold();
                            currentRun.RunProperties!.Italic = new Italic();

                            AddRun(currentRun);
                        }

                        ProcessChildrenOf(subEl, bold, true, url);

                        if (needReset)
                        {
                            currentText = new Text("");
                            currentRun = new Run(new RunProperties(), currentText);

                            if (bold) currentRun.RunProperties!.Bold = new Bold();
                            if (italic) currentRun.RunProperties!.Italic = new Italic();

                            AddRun(currentRun);
                        }
                    }
                    else if (subEl.Name == "a")
                    {
                        var neededUrl = new Uri(el.Attribute("href")!.Value);

                        if (url != neededUrl)
                        {
                            needReset = true;

                            currentText = new Text("");
                            currentRun = new Run(new RunProperties(), currentText);
                            if (bold) currentRun.RunProperties!.Bold = new Bold();
                            currentRun.RunProperties!.Italic = new Italic();
                            currentLink = new Hyperlink(currentRun)
                            {
                                Id = doc?.AllocateHyperlinkRelationship(neededUrl)
                            };

                            AddRun(currentRun);
                        }

                        ProcessChildrenOf(subEl, bold, italic, neededUrl);

                        if (needReset)
                        {
                            if (url != null)
                            {
                                currentLink = new Hyperlink
                                {
                                    Id = doc?.AllocateHyperlinkRelationship(url)
                                };
                            }

                            currentRun = new Run(new RunProperties(), currentText);

                            if (bold) currentRun.RunProperties!.Bold = new Bold();
                            if (italic) currentRun.RunProperties!.Italic = new Italic();

                            AddRun(currentRun);
                        }
                    }
                    else if (subEl.Name == "br")
                    {
                        currentText = new Text("");
                        currentRun = new Run(currentText);
                        currentParagraph = new Paragraph(currentRun);

                        paragraphs.Add(currentParagraph);
                    }
                    else
                    {
                        throw new NotImplementedException(subEl.Name.ToString());
                    }
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }
        
        ProcessChildrenOf(source, false, false, null);
        
        return paragraphs;
    }
}