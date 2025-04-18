﻿using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Formatter
{
    internal class HtmlFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string) && prefix.Equals("HTML", StringComparison.CurrentCultureIgnoreCase);
        }

        public void ApplyFormat(ITemplateProcessingContext templateContext, FormatterContext formatterContext,
            Text target)
        {
            if (formatterContext.Value is not string html || string.IsNullOrWhiteSpace(html))
            {
                target.RemoveWithEmptyParent();
                return;
            }

            // fix html - ensure starts and ends with <html>
            if (!html.StartsWith("<html>", StringComparison.CurrentCultureIgnoreCase))
            {
                html = "<html>" + html;
            }
            if (!html.EndsWith("</html>", StringComparison.CurrentCultureIgnoreCase))
            {
                html += "</html>";
            }
            var root = target.GetRoot();
            string alternativeFormatImportPartId = null;
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {
                if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                {
                    alternativeFormatImportPartId = CreateAlternativeFormatImportPart(headerPart, html);
                }
                if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                {
                    alternativeFormatImportPartId = CreateAlternativeFormatImportPart(footerPart, html);
                }
                if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                {
                    alternativeFormatImportPartId = CreateAlternativeFormatImportPart(mainDocumentPart, html);
                }
            }
            if (alternativeFormatImportPartId == null)
            {
                throw new OpenXmlTemplateException("Could not find root to insert HTML");
            }
            AltChunk altChunk = new()
            {
                Id = alternativeFormatImportPartId
            };

            var ancestorParagraph = target.GetFirstAncestor<Paragraph>();
            if (ancestorParagraph != null)
            {
                var firstPart = ancestorParagraph.SplitAfterElement(target).First();
                firstPart.InsertAfterSelf(altChunk);
            }
            else
            {
                throw new OpenXmlTemplateException("HTML import tag is not in a paragraph");
            }
            target.RemoveWithEmptyParent();
        }

        private static string CreateAlternativeFormatImportPart<T>(T parent, string html)
            where T : OpenXmlPartContainer, ISupportedRelationship<AlternativeFormatImportPart>
        {
            var alternativeFormatImportPart = parent.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
            using var memoryStream = new MemoryStream();
            using var streamWriter = new StreamWriter(memoryStream, Encoding.UTF8);
            streamWriter.Write(html);
            streamWriter.Flush();
            memoryStream.Position = 0;
            alternativeFormatImportPart.FeedData(memoryStream);
            return parent.GetIdOfPart(alternativeFormatImportPart);
        }
    }
}
