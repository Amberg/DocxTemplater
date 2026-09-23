using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocxTemplater.Formatter
{
    /// <summary>
    /// Formatter for inserting a template into the document
    /// Arguments:
    ///
    /// </summary>
    internal class SubTemplateFormatter : IFormatter
    {
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Remembers which relationship ids have already been imported per (template value, target part) so that
        // inserting the same document multiple times (e.g. in a loop) reuses the copied parts instead of
        // embedding the same image again for every insertion.
        private readonly Dictionary<(object TemplateValue, OpenXmlPart TargetPart), Dictionary<string, string>> m_importedRelationshipIds = new();

        private readonly bool m_inlineSubTemplates;

        public SubTemplateFormatter(bool inlineSubTemplates)
        {
            m_inlineSubTemplates = inlineSubTemplates;
        }

        public bool CanHandle(Type type, string prefix)
        {
            return prefix.Equals("template", StringComparison.CurrentCultureIgnoreCase) ||
                   prefix.Equals("T", StringComparison.CurrentCultureIgnoreCase);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Args.Length == 0)
            {
                throw new OpenXmlTemplateException("Template formatter requires a template name");
            }
            var loaded = LoadTemplateElements(formatterContext.Args[0]?.Trim(), templateContext.ModelLookup);
            // Keep the (optional) source document open until the merge completes, so its parts can be copied;
            // disposed at the end of the method. Null for string / OpenXmlElement templates.
            using var loadedSource = loaded.SourceDocument;
            var templateElement = loaded.Element ?? throw new OpenXmlTemplateException("Template is null or is not a valid OpenXML template");

            // Resolve the destination part up front: the table-row / -cell branches remove target's ancestor,
            // after which target.GetRoot() can no longer reach the owning part.
            var targetPart = (target.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;

            // Optional args after the template name: a selector (p/run/tr/tc) and/or "raw".
            // "raw" inserts the sub-document verbatim, WITHOUT running the template engine over it - for
            // embedding user-authored / untrusted documents whose text may legitimately contain {{ }} that
            // must survive literally (e.g. code samples or template-injection findings in a security report).
            string selector = null;
            var raw = false;
            foreach (var arg in formatterContext.Args.Skip(1))
            {
                var value = arg?.Trim();
                if (string.Equals(value, "raw", StringComparison.OrdinalIgnoreCase))
                {
                    raw = true;
                }
                else if (value is "p" or "run" or "tr" or "tc")
                {
                    selector = value;
                }
                else
                {
                    throw new OpenXmlTemplateException($"Invalid template formatter argument '{value}'");
                }
            }

            if (selector != null)
            {
                templateElement = selector switch
                {
                    "p" => templateElement.Descendants<Paragraph>().First(),
                    "run" => templateElement.Descendants<Run>().First(),
                    "tr" => templateElement.Descendants<TableRow>().First(),
                    "tc" => templateElement.Descendants<TableCell>().First(),
                    _ => throw new OpenXmlTemplateException($"Invalid selector {selector}")
                };
            }

            // Process the sub-document unless "raw" was requested. A verbatim insert skips templating, so the
            // inserted content's {{ }} tokens, loops and formatters are preserved as literal text.
            if (!raw)
            {
                // create a new Template context with replaced ModelLookup
                var templateModelLookup = new ModelLookup();
                templateModelLookup.Add("ds", formatterContext.Value);
                foreach (var models in templateContext.ModelLookup.Models.Skip(1))
                {
                    templateModelLookup.Add(models.Key, models.Value);
                }
                var variableReplacer = new VariableReplacer(templateModelLookup, templateContext.ProcessSettings);
                var scriptCompiler = new ScriptCompiler(templateModelLookup, templateContext.ProcessSettings);
                var processor = new XmlNodeTemplate(templateElement, new TemplateProcessingContext(templateContext.ProcessSettings, templateModelLookup, variableReplacer, scriptCompiler));
                processor.Process();
            }

            // Elements cloned into the target document. Their relationship references (images, external links)
            // still point at the source document's parts and must be re-imported / re-mapped below.
            var insertedElements = new List<OpenXmlElement>();

            if (m_inlineSubTemplates && templateElement is Body
                && templateElement.ChildElements is [Paragraph] or [Paragraph, SectionProperties])
            {
                templateElement = (Paragraph)templateElement.ChildElements[0];
            }

            if (templateElement is Body body)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                OpenXmlElement insertionPoint = parent.SplitAfterElement(target).First();
                foreach (var childParagaphs in body.ChildElements)
                {
                    if (childParagaphs is SectionProperties)
                    {
                        // a body-level sectPr is only valid as the last child of the body - cloning it into the
                        // middle of the target body would produce an invalid document and an unwanted section break
                        continue;
                    }
                    var clone = childParagaphs.CloneNode(true);
                    insertionPoint = insertionPoint.InsertAfterSelf(clone);
                    insertedElements.Add(clone);
                }
            }
            else if (templateElement is Paragraph paragraph)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");

                if (m_inlineSubTemplates)
                {
                    var firstRun = (OpenXmlElement)target.GetFirstAncestor<Run>() ?? throw new OpenXmlTemplateException("Could not find run to insert inline template");
                    // The split returns the part holding the target as last element - a part before it
                    // only exists if the target is not the first content of the run. Inserting before
                    // the part with the target works in both cases, anchoring on the leading part
                    // would put the template after the text that follows the tag in the same run.
                    var runWithTarget = firstRun.SplitBeforeElement(target).Last();
                    var targetProperties = runWithTarget.GetFirstChild<RunProperties>();

                    foreach (var child in paragraph.ChildElements.Where(x => x is not ParagraphProperties))
                    {
                        var clonedChild = child.CloneNode(true);
                        if (clonedChild is Run run && targetProperties != null)
                        {
                            var templateProps = run.RunProperties;

                            if (templateProps != null)
                            {
                                CascadeTargetProperties(targetProperties, templateProps);
                            }
                            else
                            {
                                run.PrependChild((RunProperties)targetProperties.CloneNode(true));
                            }
                        }
                        insertedElements.Add(runWithTarget.InsertBeforeSelf(clonedChild));
                    }
                }
                else
                {
                    var firstPart = parent.SplitAfterElement(target).First();
                    insertedElements.Add(firstPart.InsertAfterSelf(paragraph.CloneNode(true)));
                }
            }
            else if (templateElement is Run run)
            {
                var parent = target.GetFirstAncestor<Run>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                var firstPart = parent.SplitAfterElement(target).First();
                insertedElements.Add(firstPart.InsertAfterSelf(run.CloneNode(true)));
            }
            else if (templateElement is TableRow row)
            {
                var parent = target.GetFirstAncestor<TableRow>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                insertedElements.Add(parent.InsertAfterSelf(row.CloneNode(true)));
                parent.RemoveWithEmptyParent();
            }
            else if (templateElement is TableCell cell)
            {
                var parent = target.GetFirstAncestor<TableCell>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                insertedElements.Add(parent.InsertAfterSelf(cell.CloneNode(true)));
                parent.RemoveWithEmptyParent();
            }
            else
            {
                throw new OpenXmlTemplateException("Template must be a paragraph, run, table row or table cell");
            }

            // Re-import parts referenced by the inserted content (images, external links) from the source
            // document into the target part, remapping relationship ids so they resolve. Without this the
            // clones keep the source's relationship ids and the referenced media is lost.
            var sourcePart = loaded.SourceDocument?.MainDocumentPart;
            if (sourcePart != null && targetPart != null)
            {
                if (!m_importedRelationshipIds.TryGetValue((loaded.Value, targetPart), out var idMap))
                {
                    idMap = new Dictionary<string, string>(StringComparer.Ordinal);
                    m_importedRelationshipIds[(loaded.Value, targetPart)] = idMap;
                }
                ImportReferencedParts(sourcePart, targetPart, insertedElements, idMap);
            }

            target.RemoveWithEmptyParent();
        }

        private static void CascadeTargetProperties(RunProperties targetProperties, RunProperties templateProperties)
        {
            foreach (var property in targetProperties.ChildElements)
            {
                if (!templateProperties.ChildElements.Any(x => x.GetType() == property.GetType()))
                {
                    templateProperties.AddChild(property.CloneNode(true));
                }
            }
        }

        // Copies parts referenced by the inserted elements (via r:embed / r:id / r:link) from the source part
        // into the target part and rewrites the relationship ids on the clones so they resolve in the target.
        private static void ImportReferencedParts(OpenXmlPart source, OpenXmlPart target, List<OpenXmlElement> insertedElements, Dictionary<string, string> idMap)
        {
            foreach (var root in insertedElements)
            {
                foreach (var element in SelfAndDescendants(root))
                {
                    foreach (var attribute in element.GetAttributes())
                    {
                        if (attribute.NamespaceUri != RelationshipNamespace || string.IsNullOrEmpty(attribute.Value))
                        {
                            continue;
                        }
                        if (!idMap.TryGetValue(attribute.Value, out var newId))
                        {
                            newId = CopyRelationship(source, target, attribute.Value);
                            if (newId == null)
                            {
                                continue; // unresolved / unsupported relationship type - leave unchanged
                            }
                            idMap[attribute.Value] = newId;
                        }
                        element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newId));
                    }
                }
            }
        }

        private static string CopyRelationship(OpenXmlPart source, OpenXmlPart target, string relationshipId)
        {
            // Resolve the id against the source part's relationships without relying on GetPartById throwing.
            var sourcePart = source.Parts.FirstOrDefault(p => p.RelationshipId == relationshipId).OpenXmlPart;

            if (sourcePart is ImagePart image)
            {
                var newImage = AddImagePart(target, image.ContentType);
                if (newImage == null)
                {
                    return null; // target part type cannot hold an image - leave the reference unchanged
                }
                using var stream = image.GetStream();
                newImage.FeedData(stream);
                return target.GetIdOfPart(newImage);
            }

            var hyperlink = source.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId);
            if (hyperlink != null)
            {
                return target.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
            }

            var external = source.ExternalRelationships.FirstOrDefault(r => r.Id == relationshipId);
            if (external != null)
            {
                return target.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
            }

            if (sourcePart != null)
            {
                // Other embedded part types (charts, OLE objects, ...): AddPart deep-copies a part - including
                // its relationship tree - when it belongs to another package. Not every part type can be added
                // to every target part; in that case the reference is left unchanged.
                try
                {
                    return target.GetIdOfPart(target.AddPart(sourcePart));
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return null;
        }

        private static ImagePart AddImagePart(OpenXmlPart target, string contentType)
        {
            return target switch
            {
                MainDocumentPart m => m.AddImagePart(contentType),
                HeaderPart h => h.AddImagePart(contentType),
                FooterPart f => f.AddImagePart(contentType),
                _ => null
            };
        }

        private static IEnumerable<OpenXmlElement> SelfAndDescendants(OpenXmlElement element)
        {
            yield return element;
            foreach (var descendant in element.Descendants())
            {
                yield return descendant;
            }
        }

        private static LoadedTemplate LoadTemplateElements(string templateVariable, IModelLookup modelLookup)
        {
            var value = modelLookup.GetValue(templateVariable);
            if (value is string templateString)
            {
                try
                {
                    var templateNode = OpenXmlHelper.ParseOpenXmlString(templateString);
                    if (templateNode is Text text)
                    {
                        templateNode = new Run(text);
                    }

                    if (templateNode is Run run)
                    {
                        templateNode = new Paragraph(run);
                    }

                    if (templateNode is Paragraph paragraph)
                    {
                        return new LoadedTemplate(paragraph, null, null);
                    }
                    throw new OpenXmlTemplateException("String template must be a paragraph or run or text");
                }
                catch (Exception e)
                {
                    throw new OpenXmlTemplateException("Could not parse template", e);
                }
            }

            if (value is OpenXmlCompositeElement templateElement)
            {
                return new LoadedTemplate(templateElement, null, null);
            }
            if (value is byte[] byteArray)
            {
                // the fresh MemoryStream needs no disposal - key the import cache on the byte[] itself
                return OpenFromDocumentStream(new MemoryStream(byteArray), byteArray);
            }

            if (value is Stream stream)
            {
                return OpenFromDocumentStream(stream, stream);
            }
            throw new OpenXmlTemplateException("Template must be a string, OpenXmlElement, byte[] or Stream");
        }

        // The source document is kept open (returned via LoadedTemplate.SourceDocument) so its parts remain
        // readable while ImportReferencedParts copies them into the target; it is disposed by ApplyFormat.
        private static LoadedTemplate OpenFromDocumentStream(Stream stream, object templateValue)
        {
            stream.Seek(0, SeekOrigin.Begin);
            var doc = WordprocessingDocument.Open(stream, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                doc.Dispose();
                return default;
            }
            return new LoadedTemplate(body, doc, templateValue);
        }

        private readonly struct LoadedTemplate
        {
            public LoadedTemplate(OpenXmlCompositeElement element, WordprocessingDocument sourceDocument, object value)
            {
                Element = element;
                SourceDocument = sourceDocument;
                Value = value;
            }

            public OpenXmlCompositeElement Element { get; }
            public WordprocessingDocument SourceDocument { get; }
            public object Value { get; }
        }
    }
}
