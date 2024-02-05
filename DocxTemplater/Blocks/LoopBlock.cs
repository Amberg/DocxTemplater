using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class LoopBlock : ContentBlock
    {
        private readonly string m_collectionName;
        private ContentBlock m_separatorBlock;

        public LoopBlock(string collectionName, VariableReplacer variableReplacer)
            : base(variableReplacer)
        {
            m_collectionName = collectionName;
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            var model = models.GetValue(m_collectionName);
            if (model is IEnumerable<object> enumerable)
            {
                var items = enumerable.Reverse().ToList();
                int counter = 0;
                foreach (var item in items)
                {
                    counter++;
                    using var loopScope = models.OpenScope();
                    loopScope.AddVariable(m_collectionName, item);
                    var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
                    InsertContent(parentNode, cloned);
                    m_variableReplacer.ReplaceVariables(cloned);
                    ExpandChildBlocks(models, parentNode);
                    if (counter < items.Count && m_separatorBlock != null)
                    {
                        m_separatorBlock.Expand(models, parentNode);
                    }
                }
            }
            else if (model != null)
            {
                throw new OpenXmlTemplateException($"Value of {m_collectionName} is not enumerable - it is of type {model.GetType().FullName}");
            }
        }

        public void SetSeparatorBlock(ContentBlock separatorBlock)
        {
            m_separatorBlock = separatorBlock;
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}