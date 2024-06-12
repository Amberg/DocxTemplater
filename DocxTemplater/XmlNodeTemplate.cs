using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    internal class XmlNodeTemplate : TemplateProcessor
    {
        private readonly OpenXmlCompositeElement m_openXmlElement;
        private readonly MainDocumentPart m_mainDocumentPart;

        internal XmlNodeTemplate(
            OpenXmlCompositeElement openXmlElement,
            ProcessSettings settings,
            IModelLookup modelLookup,
            IVariableReplacer variableReplacer,
            IScriptCompiler scriptCompiler,
            MainDocumentPart mainDocumentPart)
            : base(settings, modelLookup, variableReplacer, scriptCompiler)
        {
            m_openXmlElement = openXmlElement;
            m_mainDocumentPart = mainDocumentPart;
        }

        public void Process()
        {
            ProcessNode(m_openXmlElement);
        }

        protected override MainDocumentPart GetMainDocumentPart()
        {
            return m_mainDocumentPart;
        }
    }
}
