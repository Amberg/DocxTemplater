using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace OpenXml.Templates
{
    internal static class ElementMarkers
    {
        public const string BeginLoop = "s";
        public const string EndLoop = "e";
        public const string Variable = "v";

        public const string MarkerAttribute = "mrk";

        public static void Mark(this OpenXmlElement element, string value)
        {
            element.SetAttribute(new OpenXmlAttribute(null, MarkerAttribute, null, value));
        }

        public static void RemoveAttribute(this OpenXmlElement element, string name)
        {
            element.RemoveAttribute(name, null);
        }

        public static bool IsMarked(this OpenXmlElement element)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == MarkerAttribute);
        }

        public static bool HasMarker(this OpenXmlElement element, string marker)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == MarkerAttribute && a.Value == marker);
        }

        public static string GetMarker(this OpenXmlElement element)
        {
            return element.ExtendedAttributes.FirstOrDefault(a => a.LocalName == MarkerAttribute).Value;
        }

        public static IEnumerable<OpenXmlElement> GetElementsWithMarker(this OpenXmlElement root, string marker)
        {
            return root.Descendants<OpenXmlElement>().Where(x => x.HasMarker(marker));
        }
    }
}
