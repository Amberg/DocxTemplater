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
            using var loadedSource = loaded.Source;
            var templateElement = loaded.Element ?? throw new OpenXmlTemplateException("Template is null or is not a valid OpenXML template");

            // Resolve the destination part up front: the table-row / -cell branches remove target's ancestor,
            // after which target.GetRoot() can no longer reach the owning part.
            var targetPart = (target.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;

            if (formatterContext.Args.Length > 1)
            {
                var selector = formatterContext.Args[1];
                templateElement = selector switch
                {
                    "p" => templateElement.Descendants<Paragraph>().First(),
                    "run" => templateElement.Descendants<Run>().First(),
                    "tr" => templateElement.Descendants<TableRow>().First(),
                    "tc" => templateElement.Descendants<TableCell>().First(),
                    _ => throw new OpenXmlTemplateException($"Invalid selector {selector}")
                };
            }

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

            // Elements cloned into the target document. Their relationship references (images, external links)
            // still point at the source document's parts and must be re-imported / re-mapped below.
            var insertedElements = new List<OpenXmlElement>();

            if (templateElement is Body body)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                OpenXmlElement insertionPoint = parent.SplitAfterElement(target).First();
                foreach (var childParagaphs in body.ChildElements)
                {
                    var clone = childParagaphs.CloneNode(true);
                    insertionPoint = insertionPoint.InsertAfterSelf(clone);
                    insertedElements.Add(clone);
                }
            }
            else if (templateElement is Paragraph paragraph)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                var firstPart = parent.SplitAfterElement(target).First();
                insertedElements.Add(firstPart.InsertAfterSelf(paragraph.CloneNode(true)));
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
            if (loaded.SourcePart != null && targetPart != null)
            {
                ImportReferencedParts(loaded.SourcePart, targetPart, insertedElements);
            }

            target.RemoveWithEmptyParent();
        }

        // Copies parts referenced by the inserted elements (via r:embed / r:id / r:link) from the source part
        // into the target part and rewrites the relationship ids on the clones so they resolve in the target.
        private static void ImportReferencedParts(OpenXmlPart source, OpenXmlPart target, List<OpenXmlElement> insertedElements)
        {
            var idMap = new Dictionary<string, string>(StringComparer.Ordinal);
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
            OpenXmlPart sourcePart = null;
            foreach (var pair in source.Parts)
            {
                if (pair.RelationshipId == relationshipId)
                {
                    sourcePart = pair.OpenXmlPart;
                    break;
                }
            }

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

            // Other embedded part types (OLE objects, charts, ...) are not copied here.
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
                var memStream = new MemoryStream(byteArray);
                return OpenFromDocumentStream(memStream, ownsStream: true);
            }

            if (value is Stream stream)
            {
                return OpenFromDocumentStream(stream, ownsStream: false);
            }
            throw new OpenXmlTemplateException("Template must be a string, OpenXmlElement, byte[] or Stream");
        }

        // The source document is kept open (returned via LoadedTemplate.Source) so its parts remain readable
        // while ImportReferencedParts copies them into the target; it is disposed by ApplyFormat.
        private static LoadedTemplate OpenFromDocumentStream(Stream stream, bool ownsStream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            WordprocessingDocument doc;
            try
            {
                doc = WordprocessingDocument.Open(stream, false);
            }
            catch
            {
                if (ownsStream)
                {
                    stream.Dispose();
                }
                throw;
            }
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                doc.Dispose();
                if (ownsStream)
                {
                    stream.Dispose();
                }
                return default;
            }
            IDisposable disposer = ownsStream ? new CompositeDisposable(doc, stream) : doc;
            return new LoadedTemplate(body, doc.MainDocumentPart, disposer);
        }

        private readonly struct LoadedTemplate
        {
            public LoadedTemplate(OpenXmlCompositeElement element, OpenXmlPart sourcePart, IDisposable source)
            {
                Element = element;
                SourcePart = sourcePart;
                Source = source;
            }

            public OpenXmlCompositeElement Element { get; }
            public OpenXmlPart SourcePart { get; }
            public IDisposable Source { get; }
        }

        private sealed class CompositeDisposable : IDisposable
        {
            private readonly IDisposable[] _items;

            public CompositeDisposable(params IDisposable[] items)
            {
                _items = items;
            }

            public void Dispose()
            {
                foreach (var item in _items)
                {
                    item?.Dispose();
                }
            }
        }
    }
}
