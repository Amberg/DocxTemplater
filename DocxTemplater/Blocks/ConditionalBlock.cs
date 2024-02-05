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
            bool conditionResult = false;
            bool removeBlock = true;
            try
            {
                conditionResult = m_scriptCompiler.CompileScript(m_condition)();
            }
            catch (OpenXmlTemplateException) when (m_scriptCompiler.ProcessSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
            {
                removeBlock = false;
            }
            if (conditionResult)
            {
                base.Expand(models, parentNode);
            }
            else
            {
                m_elseBlock?.Expand(models, parentNode);
            }

            if (removeBlock)
            {
                var element = m_insertionPoint.GetElement(parentNode);
                element.Remove();
            }

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