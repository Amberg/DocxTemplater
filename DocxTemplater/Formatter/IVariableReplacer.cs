using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace DocxTemplater.Formatter
{
    public interface IVariableReplacer
    {
        void RegisterFormatter(IFormatter formatter);

        void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content, TemplateProcessingContext templateContext);
        void ReplaceVariables(OpenXmlElement cloned, TemplateProcessingContext templateContext);
        ProcessSettings ProcessSettings { get; }
        void AddError(string errorMessage);
        void WriteErrorMessages(OpenXmlCompositeElement rootElement);
    }
}