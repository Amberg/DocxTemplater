﻿using System.Linq;
using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class InsertionPoint
    {
        private const string InsertionPointAttributeName = "IpId";
        public string Id { get; }

        private static int s_counter;

        private InsertionPoint(string id)
        {
            Id = id;
        }

        public static InsertionPoint CreateForElement(OpenXmlElement element, string name)
        {
            // element already marked return the existing insertion point
            if (element.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName))
            {
                return new InsertionPoint(element.GetAttribute(InsertionPointAttributeName, null).Value);
            }
            var insertionPoint = new InsertionPoint($"{name}_{s_counter++}");
            element.SetAttribute(new OpenXmlAttribute(null, InsertionPointAttributeName, null, insertionPoint.Id));
            return insertionPoint;
        }

        public static void RemoveAll(OpenXmlElement root)
        {
            foreach (var element in root.Descendants().Where(x => x.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName)).ToList())
            {
                if (!element.HasChildren)
                {
                    element.Remove();
                    continue;
                }
                element.RemoveAttribute(InsertionPointAttributeName, null);
            }
        }

        public static bool HasAlreadyInsertionPointMarker(OpenXmlElement element)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName);
        }

        public bool IsForElement(OpenXmlElement element)
        {
            return element.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName && a.Value == Id);
        }

        public OpenXmlElement GetElement(OpenXmlElement root)
        {
#if DEBUG
            var elements = root.Descendants<OpenXmlElement>().Where(x => x.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName && a.Value == Id)).ToList();
            if (elements.Count > 1)
            {
                throw new OpenXmlTemplateException($"Multiple elements with the same insertion point id {Id}");
            }
            return elements.FirstOrDefault();
#else
            return root.Descendants<OpenXmlElement>().FirstOrDefault(x => x.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName && a.Value == Id));
#endif
        }

        public override string ToString()
        {
            return $"IP_{Id}";
        }
    }
}
