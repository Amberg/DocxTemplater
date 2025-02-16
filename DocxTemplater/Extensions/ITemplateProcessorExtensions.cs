using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Extensions
{
    public interface ITemplateProcessorExtension
    {
        void PreProcess(OpenXmlCompositeElement content);

        void BlockContentExtracted(IReadOnlyCollection<OpenXmlElement> content);

        void ReplaceVariables(IVariableReplacer variableReplacer, IModelLookup models, OpenXmlElement parentNode, List<OpenXmlElement> newContent);
    }
}