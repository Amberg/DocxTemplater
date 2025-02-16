using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
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
        public bool CanHandle(Type type, string prefix)
        {
            return prefix.Equals("template", StringComparison.CurrentCultureIgnoreCase) ||
                   prefix.Equals("T", StringComparison.CurrentCultureIgnoreCase);
        }

        public void ApplyFormat(TemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Args.Length == 0)
            {
                throw new OpenXmlTemplateException("Template formatter requires a template name");
            }
            var templateElement = LoadTemplateElements(formatterContext.Args[0]?.Trim(), templateContext.ModelLookup) ?? throw new OpenXmlTemplateException("Template is null or is not a valid OpenXML template");

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

            if (templateElement is Body body)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                var firstPart = parent.SplitAfterElement(target).First();
                foreach (var childParagaphs in body.ChildElements)
                {
                    firstPart.InsertAfterSelf(childParagaphs.CloneNode(true));
                }
            }
            else if (templateElement is Paragraph paragraph)
            {
                var parent = target.GetFirstAncestor<Paragraph>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                var firstPart = parent.SplitAfterElement(target).First();
                firstPart.InsertAfterSelf(paragraph.CloneNode(true));
            }
            else if (templateElement is Run run)
            {
                var parent = target.GetFirstAncestor<Run>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                var firstPart = parent.SplitAfterElement(target).First();
                firstPart.InsertAfterSelf(run.CloneNode(true));
            }
            else if (templateElement is TableRow row)
            {
                var parent = target.GetFirstAncestor<TableRow>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                parent.InsertAfterSelf(row.CloneNode(true));
                parent.RemoveWithEmptyParent();
            }
            else if (templateElement is TableCell cell)
            {
                var parent = target.GetFirstAncestor<TableCell>() ?? throw new OpenXmlTemplateException("Could not find parent to insert template");
                parent.InsertAfterSelf(cell.CloneNode(true));
                parent.RemoveWithEmptyParent();
            }
            else
            {
                throw new OpenXmlTemplateException("Template must be a paragraph, run, table row or table cell");
            }
            target.RemoveWithEmptyParent();
        }

        private static OpenXmlCompositeElement LoadTemplateElements(string templateVariable, IModelLookup modelLookup)
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
                        return paragraph;
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
                return templateElement;
            }
            if (value is byte[] byteArray)
            {
                using var memStream = new MemoryStream(byteArray);
                return OpenXmlElementsFromDocumentStream(memStream);
            }

            if (value is Stream stream)
            {
                return OpenXmlElementsFromDocumentStream(stream);
            }
            throw new OpenXmlTemplateException("Template must be a string, OpenXmlElement, byte[] or Stream");
        }

        private static OpenXmlCompositeElement OpenXmlElementsFromDocumentStream(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin);
            using var doc = WordprocessingDocument.Open(stream, false);
            if (doc.MainDocumentPart == null)
            {
                return null;
            }
            var body = doc.MainDocumentPart.Document.Body;
            if (body != null)
            {
                return body;
            }
            return null;
        }
    }
}
