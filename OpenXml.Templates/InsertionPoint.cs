using DocumentFormat.OpenXml;
using System;
using System.Linq;

namespace OpenXml.Templates
{
    internal class InsertionPoint
    {
        private const string InsertionPointAttributeName = "IpId";
        public string Id { get; }

        private InsertionPoint(string id)
        {
            Id = id;
        }
        
        public static InsertionPoint CreateForElement(OpenXmlElement element, string id)
        {
            id = id ?? Guid.NewGuid().ToString("N");
            var insertionPoint = new InsertionPoint(id);
            element.SetAttribute(new OpenXmlAttribute(null, InsertionPointAttributeName, null, insertionPoint.Id));
            return insertionPoint;
        }

        public static void RemoveAll(OpenXmlElement root)
        {
            foreach (var element in root.Descendants().Where(x => x.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName)))
            {
                element.RemoveAttribute(InsertionPointAttributeName, null);
            }
        }

        public OpenXmlElement GetElement(OpenXmlElement root)
        {
            return root.Descendants<OpenXmlElement>().FirstOrDefault(x => x.ExtendedAttributes.Any(a => a.LocalName == InsertionPointAttributeName && a.Value == Id));
        }

    }
}
