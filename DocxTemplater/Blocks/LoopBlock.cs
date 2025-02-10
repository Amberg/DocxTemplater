using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocxTemplater.Blocks
{
    internal class LoopBlock : ContentBlock
    {
        private readonly string m_collectionName;

        public LoopBlock(IVariableReplacer variableReplacer, PatternType patternType, Text startTextNode, PatternMatch startMatch)
            : base(variableReplacer, patternType, startTextNode, startMatch)
        {
            m_collectionName = startMatch.Variable;
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            var model = models.GetValue(m_collectionName);
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
                throw new OpenXmlTemplateException($"Value of {m_collectionName} is not enumerable - it is of type {model.GetType().FullName}");
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}