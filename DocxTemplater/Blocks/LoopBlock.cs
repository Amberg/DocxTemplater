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

        public LoopBlock(string collectionName, VariableReplacer variableReplacer)
            : base(variableReplacer)
        {
            m_collectionName = collectionName;
        }

        public override void Expand(ModelDictionary models, OpenXmlElement parentNode)
        {
            var model = models.GetValue(m_collectionName);
            if (model is IEnumerable<object> enumerable)
            {
                int count = 0;
                foreach (var item in enumerable.Reverse())
                {
                    count++;
                    models.RemoveLoopVariable(m_collectionName);
                    models.AddLoopVariable(m_collectionName, item);

                    var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
                    InsertContent(parentNode, cloned);
                    m_variableReplacer.ReplaceVariables(cloned);
                    ExpandChildBlocks(models, parentNode);
                }
                models.RemoveLoopVariable(m_collectionName);
            }
            else
            {
                throw new OpenXmlTemplateException($"Value of {m_collectionName} is not enumerable");
            }
        }

        public override string ToString()
        {
            return $"LoopBlock: {m_collectionName}";
        }
    }
}