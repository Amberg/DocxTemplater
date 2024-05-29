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

        public override void Expand(ModelLookup models, OpenXmlElement parentNode, bool insertBeforeInsertionPoint = false)
        {
            var model = models.GetValue(m_collectionName);
            if (model is IEnumerable<object> enumerable)
            {
                var items = enumerable.Reverse().ToList();
                int counter = 0;
                foreach (var item in items)
                {
                    using var loopScope = models.OpenScope();
                    loopScope.AddVariable(m_collectionName, item);
                    var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
                    InsertContent(parentNode, cloned);

                    m_variableReplacer.ReplaceVariables(cloned);

                    ExpandChildBlocks(models, parentNode);

                    if (counter > 0 && m_separatorBlock != null)
                    {
                        m_separatorBlock.Expand(models, parentNode, true);
                    }
                    m_separatorBlock?.RemoveAnchor(parentNode);
                    counter++;
                }
            }
            else if (model != null)
            {
                throw new OpenXmlTemplateException($"Value of {m_collectionName} is not enumerable - it is of type {model.GetType().FullName}");
            }
        }

        public override void AddInnerBlock(ContentBlock block)
        {
            if (block.GetType() == typeof(ContentBlock))
            {
                m_separatorBlock = block;
            }
            else
            {
                base.AddInnerBlock(block);
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}