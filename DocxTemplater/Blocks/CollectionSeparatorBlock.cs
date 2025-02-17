using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace DocxTemplater.Blocks
{
    internal class CollectionSeparatorBlock : ContentBlock
    {
        public CollectionSeparatorBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode,
            PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            int count = (int)models.GetValue($"{ParentBlock.StartMatch.Variable.TrimStart('.')}._Idx");
            int length = (int)models.GetValue($"{ParentBlock.StartMatch.Variable.TrimStart('.')}._Length");
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
