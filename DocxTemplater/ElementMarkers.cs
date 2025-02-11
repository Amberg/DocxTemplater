using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal static class ElementMarkers
    {
        public const string MarkerAttribute = "mrk";

        public static void Mark(this OpenXmlElement element, PatternType value)
        {
            element.SetAttribute(new OpenXmlAttribute(null, MarkerAttribute, null, value.ToString()));
        }

        public static void RemoveMarker(this OpenXmlElement element)
        {
            element.RemoveAttribute(MarkerAttribute, null);
        }

        public static bool IsMarked(this OpenXmlElement element)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == MarkerAttribute);
        }

        public static bool HasMarker(this OpenXmlElement element, PatternType marker)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == MarkerAttribute && a.Value == marker.ToString());
        }

        public static PatternType GetMarker(this OpenXmlElement element)
        {
            var attribute = element.ExtendedAttributes.FirstOrDefault(a => a.LocalName == MarkerAttribute);
            return (PatternType)System.Enum.Parse<PatternType>(attribute.Value);
        }

        public static IEnumerable<OpenXmlElement> GetElementsWithMarker(this OpenXmlElement root, PatternType marker)
        {
            if (root.HasMarker(marker))
            {
                yield return root;
            }
            foreach (var element in root.Descendants())
            {
                if (element.HasMarker(marker))
                {
                    yield return element;
                }
            }
        }
    }
}
