using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OpenXml.Templates
{
    public static class OpenXmlHelper
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

        public static OpenXmlElement FindCommonParent(this OpenXmlElement element, OpenXmlElement otherElement)
        {
            if (element == null)
            {
                throw new ArgumentNullException(nameof(element));
            }
            if (otherElement == null)
            {
                throw new ArgumentNullException(nameof(otherElement));
            }
            var current = element.Parent;
            while (current != null)
            {
                if (otherElement.IsChildOf(current))
                {
                    return current;
                }
                current = current.Parent;
            }
            return null;
        }

        /// <summary>
        /// Splits the element after the given descendant element.
        /// And returns the two parts of the split element.
        /// </summary>
        public static IReadOnlyCollection<OpenXmlElement> SplitAfterElement(this OpenXmlElement elementToSplit, OpenXmlElement element)
        {
          return elementToSplit.SplitAtElement(element, false);
        }

        public static IReadOnlyCollection<OpenXmlElement> SplitBeforeElement(this OpenXmlElement elementToSplit, OpenXmlElement element)
        {
            return elementToSplit.SplitAtElement(element, true);
        }

        private static IReadOnlyCollection<OpenXmlElement> SplitAtElement(this OpenXmlElement elementToSplit, OpenXmlElement element, bool beforeElement)
        {
            var result = new List<OpenXmlElement>() { elementToSplit };
            var parent = element.Parent;
            if (parent == null)
            {
                throw new ArgumentException("cannot split a root node without parent");
            }
            var childs = beforeElement ? parent.ChildsBefore(element).ToList() : parent.ChildsAfter(element).ToList();
            if (childs.Count > 0)
            {
                var clonedParent = (OpenXmlElement)parent.CloneNode(false);
                if (parent.Parent != null)
                {
                    if (beforeElement)
                        parent.InsertBeforeSelf(clonedParent);
                    else
                        parent.InsertAfterSelf(clonedParent);
                }
                foreach (var child in childs)
                {
                    child.Remove();
                    clonedParent.AppendChild(child);
                }

                if (beforeElement)
                    result.Insert(0, clonedParent);
                else
                    result.Add(clonedParent);
            }
            if (elementToSplit == parent)
            {
                return result;
            }
            return elementToSplit.SplitAtElement(parent, beforeElement);
        }

        public static IEnumerable<OpenXmlElement> ChildsBefore(this OpenXmlElement parent, OpenXmlElement child)
        {
            foreach (var c in parent.ChildElements)
            {
                if (c == child)
                {
                    yield break;
                }
                yield return c;
            }
        }

        public static IEnumerable<OpenXmlElement> ChildsAfter(this OpenXmlElement parent, OpenXmlElement child)
        {
            bool found = false;
            foreach (var c in parent.ChildElements)
            {
                if (found)
                {
                    yield return c;
                }
                else if (c == child)
                {
                    found = true;
                }
            }
        }
        public static IEnumerable<OpenXmlElement> ChildsBetween(this OpenXmlElement parent, OpenXmlElement startChild, OpenXmlElement endChild)
        {
            bool found = false;
            foreach (var c in parent.ChildElements)
            {
                if (found)
                {
                    if (c == endChild)
                    {
                        yield break;
                    }
                    yield return c;
                }
                else if (c == startChild)
                {
                    found = true;
                }
            }
        }

        public static OpenXmlElement InsertAfterSelf(this OpenXmlElement element, IEnumerable<OpenXmlElement> elements)
        {
            foreach (var e in elements)
            {
                element.InsertAfterSelf(e);
                element = e;
            }
            return element;
        }

        public static OpenXmlElement InsertBeforeSelf(this OpenXmlElement element, IEnumerable<OpenXmlElement> elements)
        {
            foreach (var e in elements.Reverse())
            {
                element.InsertBeforeSelf(e);
            }
            return element;
        }

        public static Text MergeText(this Text first, int startIndex, Text last, int length)
        {
            var commonParent = first.FindCommonParent(last);
            if (commonParent == null)
            {
                throw new ArgumentException("Text elements must have a common parent");
            }

            if(startIndex < 0 || startIndex >= first.Text.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndex));
            }
            var matchInSameElement = first == last;
           
            if (startIndex != 0) // leading text is not part of the match
            {
                first = first.SplitAtIndex(startIndex);
            }

            if(startIndex + length < first.Text.Length) // trailing text is not part of the match in the first element
            {
                first.SplitAtIndex(length);
            }

            if(matchInSameElement)
            {
                return first;
            }

            List<OpenXmlElement> toRemove = new List<OpenXmlElement>();
            bool found = false;
            foreach (var current in commonParent.Descendants<Text>())
            {
                if (current == first)
                {
                    found = true;
                    continue;
                }
                if (!found)
                {
                    continue;
                }

                // trailing text is not part of the match int the last element
                if(first.Text.Length + current.Text.Length > length)
                {
                    var firstPart = current.Text.Substring(0, length - first.Text.Length);
                    var secondPart = current.Text.Substring(length - first.Text.Length);
                    first.Text += firstPart;
                    current.Text = secondPart;
                    break;
                }
                first.Text += current.Text;
                toRemove.Add(current);
                if (current == last)
                {
                    break;
                }
            }
            foreach (var text in toRemove)
            {
                text.RemoveWithEmptyParent();
            }
            return first;
        }

        public static Text SplitAtIndex(this Text element, int startIndexSecondPart)
        {
            if (startIndexSecondPart < 0 || startIndexSecondPart >= element.Text.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndexSecondPart));
            }
            var firstPart = element.Text.Substring(0, startIndexSecondPart);
            var secondPart = element.Text.Substring(startIndexSecondPart);
            element.Text = firstPart;
            var newElement = element.InsertAfterSelf(new Text(secondPart));
            element.Space = SpaceProcessingModeValues.Preserve;
            newElement.Space = SpaceProcessingModeValues.Preserve;
            return newElement;
        }
        
        public static string ToPrettyPrintXml(this OpenXmlElement element)
        {
            var xmldoc = XDocument.Parse("<root>"+element.InnerXml+"</root>");
            return xmldoc.ToString();
        }

        public static string PrintTree(this OpenXmlElement element, StringBuilder sb = null, int indent = 0)
        {
            if (sb == null)
            {
                sb = new StringBuilder();
            }
            sb.AppendLine($"{new string(' ', indent)}parent ({element.Parent?.GetType()?.Name}){element.GetType().Name}({element.GetType().Namespace})");
            var attributes = element.GetAttributes();
            if (attributes.Any())
            {
                sb.AppendLine($"{new string(' ', indent + 2)}Attributes:");
                foreach (var attribute in attributes)
                {
                    sb.AppendLine($"{new string(' ', indent + 4)}{attribute.LocalName} = {attribute.Value}");
                }
            }
            foreach (var child in element.ChildElements)
            {
                child.PrintTree(sb, indent + 2);
            }
            return sb.ToString();
        }


        public static OpenXmlElement GetRoot(this OpenXmlElement element)
        {
            var current = element;
            while (current.Parent != null)
            {
                current = current.Parent;
            }
            return current;
        }


        public static void RemoveWithEmptyParent(this OpenXmlElement element)
        {
            var parent = element.Parent;
            if (parent != null)
            {
                element.Remove();
                if (!parent.HasChildren)
                {
                    parent.RemoveWithEmptyParent();
                }
            }
        }
    }
}
