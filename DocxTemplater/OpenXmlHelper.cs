using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Validation;


namespace DocxTemplater
{
    public static class OpenXmlHelper
    {
        public static bool IsChildOf(this OpenXmlElement element, OpenXmlElement parent)
        {
            ArgumentNullException.ThrowIfNull(element, nameof(element));
            ArgumentNullException.ThrowIfNull(element);
            ArgumentNullException.ThrowIfNull(parent);
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

        public static Style FindTableStyleByName(this MainDocumentPart mainDocumentPart, string name)
        {
            var part = mainDocumentPart.StyleDefinitionsPart;
            if (part == null)
            {
                return null;
            }
            return part.Styles?.Elements<Style>().FirstOrDefault(x => x.StyleId == name);
        }


        /// <summary>
        /// Traverses the tree upwards and returns the first element of the given type.
        /// </summary>
        /// <typeparam name="TElement"></typeparam>
        /// <param name="element"></param>
        /// <returns></returns>
        public static OpenXmlElement ElementBeforeInDocument<TElement>(this OpenXmlElement element)
            where TElement : OpenXmlElement
        {
            var parent = element.Parent;
            while (parent != null)
            {
                var result = (parent?.Descendants<TElement>()).LastOrDefault(x => x.IsBefore(element));
                if (result != null)
                {
                    return result;
                }
                parent = parent.PreviousSibling() ?? parent.Parent;
            }
            return null;
        }

        public static TElement ElementAfterInDocument<TElement>(this OpenXmlElement element)
            where TElement : OpenXmlElement
        {
            var parent = element.Parent;
            while (parent != null)
            {
                var result = parent?.Descendants<TElement>().FirstOrDefault(x => x.IsAfter(element));
                if (result != null)
                {
                    return result;
                }
                parent = parent.NextSibling() ?? parent.Parent;
            }
            return null;
        }

        public static AbstractNum CreateNewAbstractNumbering(this Numbering numbering)
        {
            var abstractNum = new AbstractNum()
            {
                MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel }
            };
            int bulletAbstractNumId = numbering.Elements<AbstractNum>()
                .Where(x => x.AbstractNumberId != null)
                .Select(x => x.AbstractNumberId.Value).DefaultIfEmpty(0).Max() + 1;
            abstractNum.AbstractNumberId = bulletAbstractNumId;

            return numbering.InsertAfterLastChildOfSameType(abstractNum);
        }

        public static NumberingInstance CreateNewNumberingInstance(this Numbering numbering, int abstractNumberingId)
        {
            var numberIngInstance = new NumberingInstance(new AbstractNumId() { Val = abstractNumberingId })
            {
                NumberID = numbering.Elements<NumberingInstance>().Count() + 1
            };
            return numbering.InsertAfterLastChildOfSameType(numberIngInstance);
        }

        public static T InsertAfterLastChildOfSameType<T>(this OpenXmlCompositeElement parent, T newChild)
            where T : OpenXmlElement
        {
            var lastChild = parent.ChildElements.LastOrDefault(x => x.GetType() == newChild.GetType());
            if (lastChild != null)
            {
                parent.InsertAfter(newChild, lastChild);
            }
            else
            {
                parent.AppendChild(newChild);
            }
            return newChild;
        }


        public static OpenXmlElement FindCommonParent(this OpenXmlElement element, OpenXmlElement otherElement)
        {
            ArgumentNullException.ThrowIfNull(element);
            ArgumentNullException.ThrowIfNull(otherElement);
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
            if (elementToSplit == element)
            {
                return new List<OpenXmlElement>() { elementToSplit };
            }
            var result = new List<OpenXmlElement>() { elementToSplit };
            var parent = element.Parent ?? throw new ArgumentException("cannot split a root node without parent");
            var childs = beforeElement ? parent.ChildsBefore(element).ToList() : parent.ChildsAfter(element).ToList();
            if (childs.Count > 0)
            {
                var clonedParent = (OpenXmlElement)parent.CloneNode(false);
                if (parent.Parent != null)
                {
                    if (beforeElement)
                    {
                        parent.InsertBeforeSelf(clonedParent);
                    }
                    else
                    {
                        parent.InsertAfterSelf(clonedParent);
                    }
                }
                foreach (var child in childs)
                {
                    child.Remove();
                    clonedParent.AppendChild(child);
                }

                if (beforeElement)
                {
                    result.Insert(0, clonedParent);
                }
                else
                {
                    result.Add(clonedParent);
                }
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

        public static Text SplitAtIndexAndReturnLastPart(this Text element, int startIndexSecondPart)
        {
            if (startIndexSecondPart < 0 || startIndexSecondPart >= element.Text.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndexSecondPart));
            }
            var firstPart = element.Text[..startIndexSecondPart];
            var secondPart = element.Text[startIndexSecondPart..];
            element.Text = firstPart;
            var newElement = element.InsertAfterSelf(new Text(secondPart));
            element.Space = SpaceProcessingModeValues.Preserve;
            newElement.Space = SpaceProcessingModeValues.Preserve;
            return newElement;
        }

        public static string ToPrettyPrintXml(this IEnumerable<OpenXmlElement> elements)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<root>");
            foreach (var element in elements)
            {
                sb.AppendLine(element.InnerXml);
            }
            sb.AppendLine("</root>");
            var xmldoc = XDocument.Parse(sb.ToString());
            // remove all xmlns attributes
            foreach (var e in xmldoc.Descendants())
            {
                e.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            }

            foreach (var elem in xmldoc.Descendants())
            {
                elem.Name = elem.Name.LocalName;
            }
            return xmldoc.ToString();
        }

        public static string ToPrettyPrintXml(this OpenXmlElement element)
        {
            if (element == null)
            {
                return null;
            }
            var xmldoc = XDocument.Parse(element.OuterXml);
            return xmldoc.ToString();
        }

        public static string PrintTree(this OpenXmlElement element, StringBuilder sb = null, int indent = 0)
        {
            sb ??= new StringBuilder();
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

        public static T GetFirstAncestor<T>(this OpenXmlElement element)
        where T : OpenXmlElement
        {
            var current = element.Parent;
            while (current != null)
            {
                if (current is T t)
                {
                    return t;
                }
                current = current.Parent;
            }
            return null;
        }


        public static void RemoveWithEmptyParent(this OpenXmlElement element)
        {
            var parent = element.Parent;
            bool removeParent = true;
            if (parent != null)
            {
                if (element is TableCell)
                {
                    removeParent = parent.ChildElements.OfType<TableCell>().Count() == 1;
                }
                else if (element is TableRow row)
                {
                    var rowProperties = row.ChildElements.OfType<TableRowProperties>().FirstOrDefault();
                    if (rowProperties != null)
                    {
                        var nextSibling = row.NextSibling<TableRow>();
                        if (nextSibling != null)
                        {
                            var nextRowProperties =
                                nextSibling.ChildElements.OfType<TableRowProperties>().FirstOrDefault();
                            if (nextRowProperties == null)
                            {
                                rowProperties.Remove();
                                nextSibling.AddChild(rowProperties);
                            }
                        }
                    }
                }
                else
                {
                    removeParent = element.HasOnlyPropertyChildren();
                }

                if (removeParent)
                {
                    parent.RemoveChild(element);
                    RemoveWithEmptyParent(parent);
                }
            }
        }

        public static bool HasOnlyPropertyChildren(this OpenXmlElement element)
        {
            if (element == null)
            {
                return false;
            }
            return !element.ChildElements.Any(x => x is not Languages and not RunProperties and not ParagraphProperties);
        }

        public static uint GetMaxDocPropertyId(this OpenXmlPart doc)
        {
            if (doc.RootElement == null)
            {
                return 0;
            }
            return doc
                .RootElement
                .Descendants<DocProperties>()
                .Max(x => (uint?)x.Id) ?? 0;
        }

        public static void ValidateOpenXmlElement(this OpenXmlElement element)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
            var result = validator.Validate(element);
            var sb = new StringBuilder();
            sb.AppendLine("Invalid OpenXML:");
            foreach (var rInfo in result)
            {
                sb.AppendLine(rInfo.Description);
                sb.AppendLine($"Path: {rInfo.Path.XPath}");
                sb.AppendLine($"Node: {rInfo.Node} - {rInfo.RelatedNode.ToPrettyPrintXml()}");
                sb.AppendLine($"RelatedNode: {rInfo.RelatedNode} - {rInfo.RelatedNode.ToPrettyPrintXml()}");
                sb.AppendLine();
            }
            if (result.Any())
            {
                throw new OpenXmlTemplateException(sb.ToString());
            }
        }

        public static OpenXmlElement ParseOpenXmlString(string openXmlString)
        {
            XmlDocument xmlDoc = new();
            xmlDoc.LoadXml(openXmlString);
            var localName = xmlDoc.DocumentElement.LocalName;
            OpenXmlElement element = null;
            switch (localName)
            {
                case "p":
                    element = new Paragraph(openXmlString);
                    break;
                case "r":
                    element = new Run(openXmlString);
                    break;
                case "t":
                    element = new Text(openXmlString);
                    break;
                case "tbl":
                    element = new Table(openXmlString);
                    break;
                case "tr":
                    element = new TableRow(openXmlString);
                    break;
                case "tc":
                    element = new TableCell(openXmlString);
                    break;
            }
            if (element == null)
            {
                throw new InvalidOperationException("Unsupported OpenXml element: " + localName);
            }
            return element;
        }

        /* Conversion functions
         // 914 400 EMUs = 1 inch
        // 	2,54 cm = 1 inch
        // 	360 000 EMUs = 1 cm
        // we assume 96 DPI
         */
        public static int CmToEmu(int centimeter)
        {
            return centimeter * 360000;
        }

        public static int MmToEmu(int millimeter)
        {
            return millimeter * 36000;
        }

        public static int PixelsToEmu(int pixels)
        {
            return pixels * 9525;
        }

        public static int InchToEmu(int inch)
        {
            return inch * 914400;
        }


        public static int LengthToEmu(int value, string unit)
        {
            return unit switch
            {
                "cm" => CmToEmu(value),
                "mm" => MmToEmu(value),
                "px" => PixelsToEmu(value),
                "in" => InchToEmu(value),
                _ => throw new ArgumentException("Unsupported unit: " + unit)
            };
        }
    }
}
