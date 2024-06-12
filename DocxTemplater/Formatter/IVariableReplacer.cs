using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace DocxTemplater.Formatter
{
    public interface IVariableReplacer
    {
        void RegisterFormatter(IFormatter formatter);

        void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content);
        void ReplaceVariables(OpenXmlElement cloned);
    }
}