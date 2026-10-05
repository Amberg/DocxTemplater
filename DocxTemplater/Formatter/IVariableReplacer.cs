using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace DocxTemplater.Formatter
{
    public interface IVariableReplacer
    {
        void RegisterFormatter(IFormatter formatter);

        void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content, ITemplateProcessingContext templateContext);
        void ReplaceVariables(OpenXmlElement cloned, ITemplateProcessingContext templateContext);
        ProcessSettings ProcessSettings { get; }
        void AddError(string errorMessage);

        /// <summary>
        /// Adds an error to the list written into the document with <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>,
        /// formatted in the language of the <see cref="ProcessSettings"/>.
        /// </summary>
        void AddError(TemplateErrorCode errorCode, params object[] arguments);
        void WriteErrorMessages(OpenXmlCompositeElement rootElement);
    }
}