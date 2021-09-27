using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace StyleExtractor
{
    public static class StyleExtractor
    {
        public static void PrintStyles(string docPath)
        {
            XDocument styleDoc = ExtractStylesPart(docPath);
            if (styleDoc != null)
            {
                List<string> styleList = GetListOfStyles(styleDoc);
                foreach (string style in styleList)
                {
                    Console.WriteLine($"StyleName: {style}");
                }
            }
            else
            {
                Console.WriteLine("Styles returned is null!");
            }
        }

        private static List<string> GetListOfStyles(XDocument styleDoc)
        {
            List<string> styleList = new();
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace w = wordmlNamespace;

            var styles = styleDoc.Root.Elements(w + "style");
            foreach (var style in styles)
            {
                styleList.Add(style.Attribute(w + "styleId").Value);
            }
            return styleList;
        }

        public static void PrintAllStyleInfo(string docPath)
        {
            XDocument styles = ExtractStylesPart(docPath);
            if (styles != null)
                Console.WriteLine(styles);
            else
                Console.WriteLine("Styles returned is null!");
        }

        public static XDocument ExtractStylesPart(string fileName, bool getStylesWithEffectsPart = false)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            // Open the document for read access and get a reference.
            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the stylesPart variable.
                StylesPart stylesPart = null;

                if (getStylesWithEffectsPart)
                    stylesPart = docPart.StylesWithEffectsPart;
                else
                    stylesPart = docPart.StyleDefinitionsPart;

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read));
                    // Create the XDocument.
                    styles = XDocument.Load(reader);
                }
            }
            // Return the XDocument instance.
            return styles;
        }
    }
}
