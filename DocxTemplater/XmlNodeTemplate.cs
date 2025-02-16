using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    internal class XmlNodeTemplate : TemplateProcessor
    {
        private readonly OpenXmlCompositeElement m_openXmlElement;

        internal XmlNodeTemplate(OpenXmlCompositeElement openXmlElement, TemplateProcessingContext context)
            : base(context)
        {
            m_openXmlElement = openXmlElement;
        }

        public void Process()
        {
            ProcessNode(m_openXmlElement);
        }
    }
}
