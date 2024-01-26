using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Blocks
{
    internal class ConditionalBlock : ContentBlock
    {
        private readonly string m_condition;
        private readonly ScriptCompiler m_scriptCompiler;
        private IReadOnlyCollection<OpenXmlElement> m_elseContent;

        public ConditionalBlock(string condition, VariableReplacer variableReplacer, ScriptCompiler scriptCompiler)
            : base(variableReplacer)
        {
            m_condition = condition;
            m_scriptCompiler = scriptCompiler;
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            var conditionResult = m_scriptCompiler.CompileScript(m_condition)();
            var content = conditionResult ? m_content : m_elseContent;
            if (content != null)
            {
                var cloned = content.Select(x => x.CloneNode(true)).ToList();
                InsertContent(parentNode, cloned);
                m_variableReplacer.ReplaceVariables(cloned);
                ExpandChildBlocks(models, parentNode);
            }
            var element = m_leadingPart.GetElement(parentNode);
            element.Remove();
        }

        public override void SetContent(OpenXmlElement leadingPart, IReadOnlyCollection<OpenXmlElement> blockContent)
        {
            if (m_leadingPart == null)
            {
                base.SetContent(leadingPart, blockContent);
            }
            else
            {
                m_elseContent = blockContent;
            }
        }

        public override string ToString()
        {
            return $"ConditionalBlock: {m_condition}";
        }
    }
}