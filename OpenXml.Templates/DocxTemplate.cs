using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace OpenXml.Templates
{
    internal class DocxTemplate
    {
        private readonly MemoryStream m_memStream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly CharacterMap m_bodyMap;
        private readonly Dictionary<string, object> m_models;

        private readonly Regex m_regex = new Regex(@"\{\{([a-zA-Z0-9\.]+)\}(?::(\w+\(*\w*\)*))*\}", RegexOptions.Compiled);

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
            var matches = m_regex.Matches(m_bodyMap.Text);
            foreach (Match match in matches)
            {
                var variableName = match.Groups[1].Value;
                var formatter = match.Groups[2].Value;
                var value = m_models[variableName];
                m_bodyMap.ReplaceText(match.Value, value.ToString());
            }
            m_wpDocument.Save();
            m_memStream.Position = 0;
            return m_memStream;
        }
    }
}
