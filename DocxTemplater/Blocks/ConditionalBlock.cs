using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class ConditionalBlock : ContentBlock
    {
        private readonly string m_condition;
        private readonly ScriptCompiler m_scriptCompiler;

        public ConditionalBlock(string condition, VariableReplacer variableReplacer, ScriptCompiler scriptCompiler)
            : base(variableReplacer)
        {
            m_condition = condition;
            m_scriptCompiler = scriptCompiler;
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode, bool insertBeforeInsertionPoint = false)
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
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned, insertBeforeInsertionPoint);
            Debug.Assert(m_childBlocks.Count is 1 or 2);
            if (conditionResult)
            {
                m_childBlocks[0].Expand(models, parentNode);
            }
            else if (m_childBlocks.Count > 1)
            {
                m_childBlocks[1].Expand(models, parentNode);
            }
            if (removeBlock)
            {
                var element = m_insertionPoint.GetElement(parentNode);
                element.Remove();
            }
        }

        public override string ToString()
        {
            return $"ConditionalBlock: {m_condition}";
        }
    }
}