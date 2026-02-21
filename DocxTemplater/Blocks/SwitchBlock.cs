using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTemplater.Blocks
{
    internal class SwitchBlock : ContentBlock
    {
        private readonly string m_switchVariable;

        public SwitchBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
            var matchArg = startMatch.Variable.Trim();
            if (matchArg.StartsWith("switch:", StringComparison.OrdinalIgnoreCase) || matchArg.StartsWith("s:", StringComparison.OrdinalIgnoreCase))
            {
                m_switchVariable = matchArg[(matchArg.IndexOf(':') + 1)..].Trim();
            }
            else
            {
                throw new OpenXmlTemplateException($"Invalid switch block syntax: {startMatch.Variable}");
            }
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned);
            m_context.VariableReplacer.ReplaceVariables(cloned, m_context);

            var genericChildBlock = m_childBlocks.SingleOrDefault();
            if (genericChildBlock == null)
            {
                return;
            }

            CaseBlock matchedBlock = null;
            CaseBlock defaultBlock = null;

            foreach (var childBlock in genericChildBlock.ChildBlocks.OfType<CaseBlock>())
            {
                childBlock.IsMatched = false; // reset in case of repeated processing in loops
                if (childBlock.IsDefault)
                {
                    defaultBlock = childBlock;
                    continue;
                }

                if (matchedBlock == null)
                {
                    bool caseMatch = false;
                    try
                    {
                        caseMatch = m_context.ScriptCompiler.CompileScript($"{m_switchVariable} == {childBlock.MatchExpression}")();
                    }
                    catch (OpenXmlTemplateException e) when (m_context.ProcessSettings.BindingErrorHandling != BindingErrorHandling.ThrowException)
                    {
                        m_context.VariableReplacer.AddError($"{e.Message} in switch '{m_switchVariable}' case '{childBlock.MatchExpression}'");
                    }

                    if (caseMatch)
                    {
                        matchedBlock = childBlock;
                    }
                }
            }

            matchedBlock ??= defaultBlock;

            if (matchedBlock != null)
            {
                matchedBlock.IsMatched = true;
            }

            // Expanding the generic block expands its children (CaseBlocks)
            genericChildBlock.Expand(models, parentNode);

            genericChildBlock.RemoveAnchor(parentNode);
        }

        public override string ToString()
        {
            return $"SwitchBlock: {m_switchVariable}";
        }
    }
}
