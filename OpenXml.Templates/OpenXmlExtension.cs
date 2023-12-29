using System;
using DocumentFormat.OpenXml;

namespace OpenXml.Templates
{
    internal static class OpenXmlHelper
    {
        public static bool IsChildOf(this OpenXmlElement element, OpenXmlElement parent)
        {
            if (element == null)
            {
                throw new ArgumentNullException(nameof(element));
            }
            if (parent == null)
            {
                throw new ArgumentNullException(nameof(parent));
            }
            var current = element.Parent;
            while (current != null)
            {
                if (current == parent)
                {
                    return true;
                }

                current = current.Parent;
            }
            return false;
        }
    }
}
