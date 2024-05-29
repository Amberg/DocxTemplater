﻿using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    internal class XmlNodeTemplate : TemplateProcessor
    {
        private readonly OpenXmlCompositeElement m_openXmlElement;

        internal XmlNodeTemplate(OpenXmlCompositeElement openXmlElement, ProcessSettings settings, ModelLookup modelLookup, VariableReplacer variableReplacer, ScriptCompiler scriptCompiler)
            : base(settings, modelLookup, variableReplacer, scriptCompiler)
        {
            m_openXmlElement = openXmlElement;
        }

        public void Process()
        {
            ProcessNode(m_openXmlElement);
        }
    }
}
