using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace WordStyleCheck
{
    public record DiagnosticTranslationsFile(Dictionary<string, string> Translations)
    {
        public static DiagnosticTranslationsFile LoadFromDocx(string path)
        {
            var translations = new Dictionary<string, string>();
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                string part = "";
                string currentText = "";
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
                                translations[part] = currentText;
                                currentText = "";
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
                        currentText += el.OuterXml;
                    }
                }
            }

            return new DiagnosticTranslationsFile(translations);
        } 
    }
}
