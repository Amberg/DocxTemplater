using DocumentFormat.OpenXml;

namespace DocxTemplater
{
    public interface ITemplateProcessorExtension
	{
        void PreProcess(OpenXmlCompositeElement content);
    }
}
