using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace DocxTemplater.Formatter
{
    public interface IVariableReplacer
    {
        void RegisterFormatter(IFormatter formatter);

        void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content);
        void ReplaceVariables(OpenXmlElement cloned);
        ProcessSettings ProcessSettings { get; }
        void AddError(string errorMessage);
        void WriteErrorMessages(OpenXmlCompositeElement rootElement);
    }
}