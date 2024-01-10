using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;
using System.Collections.Generic;

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

        public override void Expand(ModelDictionary models, OpenXmlElement parentNode)
        {
            var conditionResult = m_scriptCompiler.CompileScript(m_condition)();
            var content = conditionResult ? m_content : m_elseContent;
            if (content != null)
            {
                var paragraphs = CreateBlockContentForCurrentVariableStack(content);
                InsertContent(parentNode, paragraphs);
                ExpandChildBlocks(models, parentNode);
            }
            var element = m_leadingPart.GetElement(parentNode);
            element.Remove();
        }

        public override void SetContent(OpenXmlElement leadingPart, IReadOnlyCollection<OpenXmlElement> loopContent)
        {
            if (m_leadingPart == null)
            {
                base.SetContent(leadingPart, loopContent);
            }
            else
            {
                m_elseContent = loopContent;
            }
        }

        public override string ToString()
        {
            return $"ConditionalBlock: {m_condition}";
        }
    }
}