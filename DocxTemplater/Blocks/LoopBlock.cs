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
        private IReadOnlyCollection<OpenXmlElement> m_separatorBlock;

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
                        var clonedSeparator = m_separatorBlock.Select(x => x.CloneNode(true)).ToList();
                        InsertContent(parentNode, clonedSeparator);
                        m_variableReplacer.ReplaceVariables(clonedSeparator);
                        ExpandChildBlocks(models, parentNode);
                    }
                }
            }
            else
            {
                throw new OpenXmlTemplateException($"Value of {m_collectionName} is not enumerable");
            }
        }

        public override void SetContent(OpenXmlElement leadingPart, IReadOnlyCollection<OpenXmlElement> blockContent)
        {
            if (m_leadingPart == null)
            {
                base.SetContent(leadingPart, blockContent);
            }
            else
            {
                m_separatorBlock = blockContent;
                leadingPart.RemoveWithEmptyParent();
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}