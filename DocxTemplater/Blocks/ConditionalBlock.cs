using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class ConditionalBlock : ContentBlock
    {
        private readonly string m_condition;
        private readonly ScriptCompiler m_scriptCompiler;
        private ContentBlock m_elseBlock;

        public ConditionalBlock(string condition, VariableReplacer variableReplacer, ScriptCompiler scriptCompiler)
            : base(variableReplacer)
        {
            m_condition = condition;
            m_scriptCompiler = scriptCompiler;
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            // todo; catch script errors here and report them or keep element in document
            var conditionResult = m_scriptCompiler.CompileScript(m_condition)();
            if (conditionResult)
            {
                base.Expand(models, parentNode);
            }
            else
            {
                m_elseBlock.Expand(models, parentNode);
            }
            var element = m_insertionPoint.GetElement(parentNode);
            element.Remove();
        }

        public void SetElseBlock(ContentBlock elseBlock)
        {
            m_elseBlock = elseBlock;
        }

        public override string ToString()
        {
            return $"ConditionalBlock: {m_condition}";
        }
    }
}