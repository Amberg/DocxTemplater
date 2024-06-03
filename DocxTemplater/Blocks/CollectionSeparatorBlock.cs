using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using System.Collections.Generic;

namespace DocxTemplater.Blocks
{
    internal class CollectionSeparatorBlock : ContentBlock
    {
        public CollectionSeparatorBlock(VariableReplacer variableReplacer, PatternType patternType, Text startTextNode,
            PatternMatch startMatch)
            : base(variableReplacer, patternType, startTextNode, startMatch)
        {
        }

        public override void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            int count = (int)models.GetValue($"{ParentBlock.StartMatch.Variable}._Count");
            int length = (int)models.GetValue($"{ParentBlock.StartMatch.Variable}._Length");
            // last element is rendered first - get length and count ot to not render the last separator
            if (length - count == 0)
            {
                return;
            }
            base.Expand(models, parentNode);
        }

        protected override void InsertContent(OpenXmlElement parentNode, IEnumerable<OpenXmlElement> paragraphs)
        {
            var element = m_insertionPoint.GetElement(parentNode) ?? throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            element.InsertBeforeSelf(paragraphs);
        }
    }
}
