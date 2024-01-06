﻿using DocumentFormat.OpenXml.Wordprocessing;
using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using System.Text;

namespace OpenXml.Templates.Formatter
{
    internal class HtmlFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(string) && prefix.Equals("HTML", StringComparison.CurrentCultureIgnoreCase);
        }

        public void ApplyFormat(string modelPath, object value, string prefix, string[] args, Text target)
        {
            if (value is not string html)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(html))
            {
                return;
            }

            var root = target.GetRoot();
            string alternativeFormatImportPartId = null;
            if (root is OpenXmlPartRootElement openXmlPartRootElement && openXmlPartRootElement.OpenXmlPart != null)
            {
                if (openXmlPartRootElement.OpenXmlPart is HeaderPart headerPart)
                {
                    alternativeFormatImportPartId = HtmlFormatter.CreateAlternativeFormatImportPart(headerPart, html);
                }
                if (openXmlPartRootElement.OpenXmlPart is FooterPart footerPart)
                {
                    alternativeFormatImportPartId = HtmlFormatter.CreateAlternativeFormatImportPart(footerPart, html);
                }
                if (openXmlPartRootElement.OpenXmlPart is MainDocumentPart mainDocumentPart)
                {
                    alternativeFormatImportPartId = HtmlFormatter.CreateAlternativeFormatImportPart(mainDocumentPart, html);
                }
            }
            if (alternativeFormatImportPartId == null)
            {
                throw new OpenXmlTemplateException("Could not find a valid image part");
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
