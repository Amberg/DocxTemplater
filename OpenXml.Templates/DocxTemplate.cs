﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Templates.Formatter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenXml.Templates
{
    public class DocxTemplate : IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly ModelDictionary m_models;

        private static readonly Regex patternRegex = new (@"\{\{([#/])?([a-zA-Z0-9\.]+)\}(?::(\w+\(*\w*\)*))?\}", RegexOptions.Compiled);
        private readonly VariableReplacer m_variableReplacer;

        public DocxTemplate(Stream docXStream)
        {
            m_stream = new MemoryStream();
            docXStream.CopyTo(m_stream);
            m_stream.Position = 0;
            m_wpDocument = WordprocessingDocument.Open(m_stream, true);
            m_models = new ModelDictionary();
            m_variableReplacer = new VariableReplacer(m_models);
        }

        public void RegisterFormatter(IFormatter formatter)
        {
            m_variableReplacer.RegisterFormatter(formatter);
        }

        public void AddModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public Stream Process()
        { 
            m_models.SetModelPrefix();

            ProcessNode(m_wpDocument.MainDocumentPart.HeaderParts.First().Header);

            Console.WriteLine("----------- Header --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.HeaderParts.First().Header.ToPrettyPrintXml());
            Console.WriteLine("-----------  --------");
            Console.WriteLine("-----------  --------");

            ProcessNode(m_wpDocument.MainDocumentPart.Document.Body);
            m_wpDocument.Save();
            m_stream.Position = 0;
            return m_stream;
        }

        private void ProcessNode(OpenXmlCompositeElement content)
        {


#if DEBUG
            Console.WriteLine("----------- Original --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
#endif

            IsolateAndMergeTextTemplateMarkers(content);

#if DEBUG
            Console.WriteLine("----------- Isolate Texts --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
#endif

            var loops = ExpandLoops(content);

            m_variableReplacer.ReplaceVariables(content);

#if DEBUG
            Console.WriteLine("----------- After Loops --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
#endif
            foreach (var loop in loops)
            {
                loop.Expand(m_models, m_wpDocument.MainDocumentPart.Document.Body);
            }
            Cleanup(m_wpDocument.MainDocumentPart.Document.Body);
#if DEBUG
            Console.WriteLine("----------- Completed --------");
            Console.WriteLine(m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
#endif
        }

        private void IsolateAndMergeTextTemplateMarkers(OpenXmlCompositeElement content)
        {
            var charMap = new CharacterMap(content);
            foreach (Match m in patternRegex.Matches(charMap.Text))
            {
                var firstChar = charMap[m.Index];
                var lastChar = charMap[m.Index + m.Length - 1];
                var firstText = (Text)firstChar.Element;
                var lastText = (Text)lastChar.Element;
                var mergedText = firstText.MergeText(firstChar.Index, lastText, m.Length);
                if (m.Groups[1].Value == "#")
                {
                    mergedText.Mark(ElementMarkers.BeginLoop);
                }
                else if (m.Groups[1].Value == "/")
                {
                    mergedText.Mark(ElementMarkers.EndLoop);
                }
                else
                {
                    mergedText.Mark(ElementMarkers.Variable);
                }
                // TODO: Ist this posible without recreate charMap?
                charMap.MarkAsDirty();
            }
        }

        private void Cleanup(OpenXmlCompositeElement element)
        {
          InsertionPoint.RemoveAll(element);
          foreach (var emptyParagraph in element.Descendants<Text>().Where(x => x.IsMarked()).ToList())
          {
              var value = emptyParagraph.GetMarker();
              if (value == ElementMarkers.BeginLoop || value == ElementMarkers.EndLoop)
              {
                  emptyParagraph.RemoveWithEmptyParent();
              }
              else
              {
                  emptyParagraph.RemoveAttribute("mrk", null);
              }
          }
        }

        private IReadOnlyCollection<LoopBlock> ExpandLoops(OpenXmlCompositeElement element)
        {
            var collectionStack = new Stack<(string Name, Text startText, List<LoopBlock> InnerBlocks)>();
            collectionStack.Push(("Root", null, new List<LoopBlock>()));
            // find all begin or end markers
            foreach (var text in element.Descendants<Text>().Where(x => x.IsMarked()))
            {
                var value = text.GetMarker();
                if (value == ElementMarkers.BeginLoop)
                {
                    var matches = patternRegex.Matches(text.Text);
                    var variableName = matches[0].Groups[2].Value;
                    collectionStack.Push((variableName, text, new List<LoopBlock>()));
                }

                if (value == ElementMarkers.EndLoop)
                {
                    var matches = patternRegex.Matches(text.Text);
                    var variableName = matches[0].Groups[2].Value;
                    var enumerationData = collectionStack.Pop();
                    if (enumerationData.Name != variableName) 
                    {
                        throw new OpenXmlTemplateException($"Collection {enumerationData.Name} is not closed");
                    }
                    var nodesInLoop = ExtractLoopContent(enumerationData.startText, text, out var leadingPart);
                    m_models.Remove(enumerationData.Name);
                    collectionStack.Peek().InnerBlocks.Add(new LoopBlock(InsertionPoint.CreateForElement(leadingPart, enumerationData.Name), nodesInLoop, enumerationData.InnerBlocks, enumerationData.Name,this, m_variableReplacer));
                }
            }
            var root = collectionStack.Pop();
            return root.InnerBlocks;
        }


        internal IReadOnlyCollection<OpenXmlElement> ExtractLoopContent(OpenXmlElement startText, OpenXmlElement endText, out OpenXmlElement leadingPart)
        {
            var commonParent = startText.FindCommonParent(endText);
            if (commonParent == null)
            {
                throw new OpenXmlTemplateException("Start and end text are not in the same tree");
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

        private class LoopBlock
        {
            private readonly InsertionPoint m_leadingPart;
            private readonly IReadOnlyCollection<OpenXmlElement> m_loopContent;
            private readonly IEnumerable<LoopBlock> m_childBlocks;
            private readonly string m_name;
            private readonly DocxTemplate m_doc;
            private readonly VariableReplacer m_variableReplacer;

            public LoopBlock(InsertionPoint insertionPoint, IReadOnlyCollection<OpenXmlElement> loopContent, IEnumerable<LoopBlock> childBlocks, string name, DocxTemplate doc, VariableReplacer variableReplacer)
            {
                m_leadingPart = insertionPoint;
                m_loopContent = loopContent;
                m_childBlocks = childBlocks;
                m_name = name;
                m_doc = doc;
                m_variableReplacer = variableReplacer;
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
                            m_variableReplacer.ReplaceVariables(cloned);
                            return cloned;
                        });


                        // insert
                        var element = insertionMPoint.GetElement(parentNode);
                        if (element == null)
                        {
                            Console.WriteLine(parentNode.ToPrettyPrintXml());
                            throw new OpenXmlTemplateException($"Insertion point {insertionMPoint.Id} not found");
                        }
                        element.InsertAfterSelf(paragraphs);
#if DEBUG
                        Console.WriteLine("----------- After Loop --------");
                        Console.WriteLine(m_doc.m_wpDocument.MainDocumentPart.Document.ToPrettyPrintXml());
#endif

                        foreach (var child in m_childBlocks)
                        {
                            child.Expand(models, parentNode);
                        }
                    }
                    models.Remove(m_name);
                }
                else
                {
                   throw new OpenXmlTemplateException($"Value of {m_name} is not enumerable");
                }
            }
        }
    }
}
