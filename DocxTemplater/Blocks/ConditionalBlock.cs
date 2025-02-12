using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class ConditionalBlock : ContentBlock
    {
        private readonly string m_condition;
        private readonly IScriptCompiler m_scriptCompiler;


        public ConditionalBlock(IVariableReplacer variableReplacer, IScriptCompiler scriptCompiler, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(variableReplacer, patternType, startTextNode, startMatch)
        {
            m_condition = startMatch.Condition;
            m_scriptCompiler = scriptCompiler;
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            bool conditionResult = false;
            try
            {
                conditionResult = m_scriptCompiler.CompileScript(m_condition)();
            }
            catch (OpenXmlTemplateException e) when (m_scriptCompiler.ProcessSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
            {
                m_variableReplacer.AddError($"{e.Message} in condition '{m_condition}'");
            }
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned);
            m_variableReplacer.ReplaceVariables(cloned);
            Debug.Assert(m_childBlocks.Count is 1 or 2);

            var elseBlock = m_childBlocks.Count > 1 ? m_childBlocks[1] : null;
            var conditionBlock = m_childBlocks[0];
            if (conditionResult)
            {
                conditionBlock.Expand(models, parentNode);
            }
            else if (m_childBlocks.Count > 1)
            {
                elseBlock?.Expand(models, parentNode);
            }
            conditionBlock.RemoveAnchor(parentNode);
            elseBlock?.RemoveAnchor(parentNode);
        }

        public override string ToString()
        {
            return $"ConditionalBlock: {m_condition}";
        }
    }
}