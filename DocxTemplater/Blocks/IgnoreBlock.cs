using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Schema;
using System.Linq;

namespace DocxTemplater.Blocks
{
    internal class IgnoreBlock : ContentBlock
    {
        public IgnoreBlock(ITemplateProcessingContext context, PatternType patternType, Text startTextNode,
            PatternMatch startMatch)
            : base(context, patternType, startTextNode, startMatch)
        {
        }

        public override void Expand(IModelLookup models, OpenXmlElement parentNode)
        {
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned); ;
        }

        public override void ExtractContentRecursively()
        {
            m_content = ParentNode.ChildsBetween(FirstElement, LastElement).ToList();
            foreach (var content in m_content)
            {
                content.Remove();
            }
        }

        public override void Validate()
        {

        }

        public override string ToString()
        {
            return $"IgnoreBlock";
        }

        public override void CollectSchema(SchemaBuilder builder)
        {
            // Content inside :ignore is dropped at render time; do not contribute to the schema.
        }
    }
}
