using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace DocxTemplater.Extensions
{
    public interface ITemplateProcessorExtension
    {
        void PreProcess(OpenXmlCompositeElement content);

        void ReplaceVariables(ITemplateProcessingContext templateContext, OpenXmlElement parentNode, List<OpenXmlElement> newContent);
    }
}