using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenXml.Templates
{
    internal class DocxTemplate : IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly CharacterMap m_bodyMap;
        private readonly ModelDictionary m_models;

        private static readonly Regex m_collectionRegex = new (@"\{\{([#/])([a-zA-Z0-9\.]+)\}\}", RegexOptions.Compiled);
        public DocxTemplate(Stream docXStream)
        {
            m_stream = new MemoryStream();
            docXStream.CopyTo(m_stream);
            m_stream.Position = 0;
            m_wpDocument = WordprocessingDocument.Open(m_stream, true);
            m_bodyMap = new CharacterMap(m_wpDocument.MainDocumentPart.Document.Body);
            m_models = new ModelDictionary();
        }

        public void AddModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public Stream Process()
        {
            m_models.SetModelPrefix();
            var loops = ExpandLoops(m_bodyMap);
            m_bodyMap.ReplaceVariables(m_models);
            Console.WriteLine("----------- After Loops --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
            foreach (var loop in loops)
            {
                loop.Expand(m_models, m_wpDocument.MainDocumentPart.Document.Body);
            }
            Cleanup(m_wpDocument.MainDocumentPart.Document.Body);
            Console.WriteLine("----------- Completed --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
            m_wpDocument.Save();
            m_stream.Position = 0;
            return m_stream;
        }

        private void Cleanup(OpenXmlCompositeElement element)
        {
          InsertionPoint.RemoveAll(element);
          foreach (var emptyParagraph in element.Descendants<Text>().Where(x => x.Text.StartsWith("{{#") || x.Text.StartsWith("{{/")).ToList())
          {
              emptyParagraph.RemoveWithEmptyParent();
          }
          ////foreach (var emptyParagraph in element.Descendants<Paragraph>().Where(x => x.Descendants().All(c => c is not Text)).ToList())
          ////{
          ////    emptyParagraph.RemoveWithEmptyParent();
          ////}
          ////foreach (var emptyParagraph in element.Descendants<TableRow>().Where(x => x.Descendants().All(c => c is not Text)).ToList())
          ////{
          ////    emptyParagraph.RemoveWithEmptyParent();
          ////}
          ////foreach (var emptyParagraph in element.Descendants<Run>().Where(x => x.Descendants().All(c => c is not Text)).ToList())
          ////{
          ////    emptyParagraph.RemoveWithEmptyParent();
          ////}
        }

        private IReadOnlyCollection<LoopBlocks> ExpandLoops(CharacterMap characterMap)
        {
            var matches = ((IReadOnlyCollection<Match>)m_collectionRegex.Matches(characterMap.Text)).Select(m =>
            {
                
                var firstChar = characterMap[m.Index];
                var lastChar = characterMap[m.Index + m.Length - 1];
                var firstText = (Text)firstChar.Element;
                var lastText = (Text)lastChar.Element;

                firstText.MergeText(firstChar.Index, lastText, m.Length);

                return new
                {
                    Textz = firstText,
                    Prefix = m.Groups[1].Value,
                    VariableName = m.Groups[2].Value,
                    FirstCharacter = m_bodyMap[m.Index],
                    LastCharacter = m_bodyMap[m.Index + m.Length - 1],
                };
            }).ToList();


            m_bodyMap.MarkAsDirty();

            var collectionStack = new Stack<(string Name, Text startText, List<LoopBlocks> InnerBlocks)>();
            collectionStack.Push(("Root", null, new List<LoopBlocks>()));
            foreach (var m in matches)
            {
                if (m.Prefix == "#")
                { 
                   collectionStack.Push((m.VariableName, m.Textz, new List<LoopBlocks>()));
                }
                else if (m.Prefix == "/")
                {
                    var enumerationData = collectionStack.Pop();
                    if (enumerationData.Name != m.VariableName)
                    {
                        throw new Exception($"Collection {enumerationData.Name} is not closed");
                    }
                    var nodesInLoop = ExtractLoopContent(enumerationData.startText, m.Textz, out var leadingPart);
                    m_bodyMap.MarkAsDirty();
                    collectionStack.Peek().InnerBlocks.Add(new LoopBlocks(InsertionPoint.CreateForElement(leadingPart, enumerationData.Name), nodesInLoop, enumerationData.InnerBlocks, enumerationData.Name, this));
                }
            }
            var root = collectionStack.Pop();
            return root.InnerBlocks;
        }


        internal IReadOnlyCollection<OpenXmlElement> ExtractLoopContent(OpenXmlElement startText, OpenXmlElement endText, out OpenXmlElement leadingPart)
        {
            // TODO: handle start marker in same run;
            var commonParent = startText.FindCommonParent(endText);
            if (commonParent == null)
            {
                throw new Exception("Start and end text are not in the same tree");
            }
            var result = new List<OpenXmlElement>();


            if (commonParent is TableRow)
            {
                var previousRow = commonParent.PreviousSibling();
                if (previousRow == null)
                {
                    commonParent.InsertBeforeSelf(new TableRow());
                }
                leadingPart = commonParent.PreviousSibling();
                commonParent.Remove();
                result.Add(commonParent);
            }
            else
            {
                // find childs of commmon parent that contains start and end text
                var startChildOfCommonParent = commonParent.ChildElements.Single(c =>
                    c == startText || c.Descendants<Text>().Any(d => d == startText));
                var endChildOfCommonParent =
                    commonParent.ChildElements.Single(c =>
                        c == endText || c.Descendants<Text>().Any(d => d == endText));

                var startSplit = startChildOfCommonParent.SplitAfterElement(startText);
                leadingPart = startSplit.First();
                if (startChildOfCommonParent == endChildOfCommonParent)
                {
                    result.AddRange(commonParent.ChildsBetween(startSplit.First(), endChildOfCommonParent).ToList());
                }
                else
                {
                    var endSplit = endChildOfCommonParent.SplitBeforeElement(endText);
                    result.AddRange(commonParent.ChildsBetween(leadingPart, endSplit.Last()).ToList());
                }

                foreach (var element in result)
                {
                    element.Remove();
                }
            }
            return result;
        }


        public void Dispose()
        {
            m_stream?.Dispose();
            m_wpDocument?.Dispose();
        }

        private class LoopBlocks
        {
            private readonly InsertionPoint m_leadingPart;
            private readonly IReadOnlyCollection<OpenXmlElement> m_loopContent;
            private readonly IEnumerable<LoopBlocks> m_childBlocks;
            private readonly string m_name;
            private readonly DocxTemplate m_doc;

            public LoopBlocks(InsertionPoint insertionPoint, IReadOnlyCollection<OpenXmlElement> loopContent, IEnumerable<LoopBlocks> childBlocks, string name, DocxTemplate doc)
            {
                m_leadingPart = insertionPoint;
                m_loopContent = loopContent;
                m_childBlocks = childBlocks;
                m_name = name;
                m_doc = doc;
            }

            public void Expand(ModelDictionary models, OpenXmlElement parentNode)
            {
                var insertionMPoint = m_leadingPart;
                var model = models.GetValue(m_name);
                if(model is IEnumerable<object> enumerable)
                {
                    int count = 0;
                    foreach (var item in enumerable.Reverse())
                    {
                        count++;
                        models.Remove(m_name);
                        models.Add(m_name, item);
                        
                        var paragraphs = m_loopContent.Select(x =>
                        {
                            var cloned = (OpenXmlCompositeElement)x.CloneNode(true);
                            var charMap = new CharacterMap(cloned);
                            charMap.ReplaceVariables(models);
                            return cloned;
                        });


                        // insert
                        var element = insertionMPoint.GetElement(parentNode);
                        if (element == null)
                        {
                            Console.WriteLine(parentNode.ToPrettyPrintXml());
                            throw new Exception($"Insertion point {insertionMPoint.Id} not found");
                        }
                        element.InsertAfterSelf(paragraphs);

                        Console.WriteLine("----------- After Loop --------");
                        Console.WriteLine(m_doc.m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());

                        foreach (var child in m_childBlocks)
                        {
                            child.Expand(models, parentNode);
                        }
                    }
                    models.Remove(m_name);
                }
                else
                {
                   throw new Exception($"Value of {m_name} is not enumerable");
                }
            }
        }
    }
}
