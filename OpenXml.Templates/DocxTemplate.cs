using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenXml.Templates
{
    internal class DocxTemplate
    {
        private readonly MemoryStream m_memStream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly CharacterMap m_bodyMap;
        private readonly Dictionary<string, object> m_models;


        private readonly Regex m_collectionRegex = new (@"\{\{([#/])([a-zA-Z0-9\.]+)\}\}", RegexOptions.Compiled);
        private readonly Regex m_regex = new (@"\{\{(#)*([a-zA-Z0-9\.]+)\}(?::(\w+\(*\w*\)*))*\}", RegexOptions.Compiled);

        public DocxTemplate(Stream docXStream)
        {

            m_memStream = new MemoryStream();
            docXStream.CopyTo(m_memStream);
            m_memStream.Position = 0;
            m_wpDocument = WordprocessingDocument.Open(m_memStream, true);
            m_bodyMap = new CharacterMap(m_wpDocument.MainDocumentPart.Document.Body);
            m_models = new Dictionary<string, object>();
        }

        public void AddModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public Stream Process()
        {
            ExpandLoops(m_bodyMap);
            ReplaceVariables(m_bodyMap);

            m_wpDocument.Save();
            m_memStream.Position = 0;
            return m_memStream;
        }

        private void ExpandLoops(CharacterMap characterMap)
        {
            var matches = ((IReadOnlyCollection<Match>)m_collectionRegex.Matches(characterMap.Text)).Select(m => new
            {
                Prefix = m.Groups[1].Value,
                VariableName = m.Groups[2].Value,
                Text = (Text)m_bodyMap[m.Index].Element,
                MatchText = m.Value
            }).ToList();
            var collectionStack = new Stack<(string Name, Text startText, int count)>();
            foreach (var m in matches)
            {
                if (m.Prefix == "#")
                {
                    var value = GetValue(m.VariableName);
                    if (value is ICollection enumerable)
                    {
                        collectionStack.Push((m.VariableName, m.Text, enumerable.Count));
                        m.Text.Text = m.Text.Text.Replace(m.MatchText, string.Empty);
                    }
                    else
                    {
                        throw new Exception($"Value of {m.VariableName} is not enumerable");
                    }
                }
                else if (m.Prefix == "/")
                {
                    var enumerationData = collectionStack.Pop();
                    if (enumerationData.Name != m.VariableName)
                    {
                        throw new Exception($"Collection {enumerationData.Name} is not closed");
                    }
                    m.Text.Text = m.Text.Text.Replace(m.MatchText, string.Empty);
                    // get all text between collection start and collection end
                    var endText = m.Text;
                    var nodeInsideLoop = m_bodyMap.CutBetween(enumerationData.startText, endText);
                    OpenXmlElement startElement = enumerationData.startText;
                    startElement.InsertBeforeSelf(new Text($"{{{{#{m.VariableName}}}}}"));
                    for (int i = 0; i < enumerationData.count; i++)
                    {
                        startElement = m_bodyMap.InsertParagraphsAfterText(startElement, nodeInsideLoop.Select(x => x.CloneNode(true)));
                    }
                }
            }
        }

        private void ReplaceVariables(CharacterMap characterMap)
        {
            var textMatches = ((IReadOnlyCollection<Match>)m_regex.Matches(characterMap.Text)).Select(m => new
            {
                Prefix = m.Groups[1].Value,
                VariableName = m.Groups[2].Value,
                Format = m.Groups[3].Value,
                Text = (Text)m_bodyMap[m.Index].Element,
                MatchText = m.Value
            }).ToList();
            foreach (var m in textMatches)
            {
                if (m.Prefix == "#")
                {
                    var value = GetValue(m.VariableName);
                    if (value is IEnumerator)
                    {
                        m_models.Remove(m.VariableName);
                        value = GetValue(m.VariableName);
                    }
                    if (value is IEnumerable enumerable)
                    {
                        var enumerator = enumerable.GetEnumerator();
                        enumerator.MoveNext();
                        m_models.Add(m.VariableName, enumerator);
                        m.Text.Text = m.Text.Text.Replace(m.MatchText, string.Empty);
                    }
                    else
                    {
                        throw new Exception($"Value of {m.VariableName} is not enumerable");
                    }
                }
                else
                {
                    var value = GetValue(m.VariableName);
                    if (value != null)
                    {
                        m.Text.Text = m.Text.Text.Replace(m.MatchText, value.ToString());
                    }
                }
            }
        }

        private object GetValue(string variableName)
        {
            var parts = variableName.Split('.');
            var path = parts[0];
            object model = m_models[path];
            for (int i = 1; i < parts.Length; i++)
            {
                path += $".{parts[i]}";
                if (!m_models.TryGetValue(path, out var nextModell))
                {
                    var property = model.GetType().GetProperty(parts[i]);
                    if (property != null)
                    {
                        model = property.GetValue(model);
                    }
                }
                else
                {
                    model = nextModell;
                }
                // if variable name is a collection, return enumerator otherwise use the current value of  the enumerator to get the next value
                if(model is IEnumerator enumerator && i + 1 < parts.Length)
                {
                  model = enumerator.Current;
                }
            }
            return model;
        }
    }
}
