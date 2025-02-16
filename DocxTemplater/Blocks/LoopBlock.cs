using System;
using System.Collections;
using System.Linq;
using DocumentFormat.OpenXml;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Blocks
{
    internal class LoopBlock : ContentBlock
    {
        private readonly string m_collectionName;

        public LoopBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
            m_collectionName = startMatch.Variable;
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            object model = null;
            try
            {
                model = models.GetValue(m_collectionName);
            }
            catch (OpenXmlTemplateException e) when (m_context.ProcessSettings.BindingErrorHandling !=
                                                     BindingErrorHandling.ThrowException)
            {
                if (m_context.ProcessSettings.BindingErrorHandling ==
                    BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_context.VariableReplacer.AddError(e.Message);
                }
            }

            if (model is IEnumerable enumerable)
            {
                var items = enumerable.Cast<object>().Reverse().ToList();
                int counter = items.Count;
                foreach (var item in items)
                {
                    using var loopScope = models.OpenScope();
                    loopScope.AddVariable(m_collectionName, item);
                    loopScope.AddVariable($"{m_collectionName}._Idx", counter);
                    loopScope.AddVariable($"{m_collectionName}._Length", items.Count);
                    base.Expand(models, parentNode);
                    counter--;
                }
            }
            else if (model != null)
            {
                if (m_context.VariableReplacer.ProcessSettings.BindingErrorHandling == BindingErrorHandling.ThrowException)
                {
                    throw new OpenXmlTemplateException(
                        $"'{m_collectionName}' is not enumerable - it is of type {model.GetType().FullName}");
                }
                else
                {
                    m_context.VariableReplacer.AddError(
                        $"'{m_collectionName}' is not enumerable - it is of type {model.GetType().FullName}");
                }
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}