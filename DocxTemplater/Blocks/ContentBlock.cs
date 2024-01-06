using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class ContentBlock
    {
        protected InsertionPoint m_leadingPart;
        protected IReadOnlyCollection<OpenXmlElement> m_content;
        protected readonly List<ContentBlock> m_childBlocks;
        protected readonly VariableReplacer m_variableReplacer;

        public ContentBlock(VariableReplacer variableReplacer)
        {
            m_leadingPart = null;
            m_content = new List<OpenXmlElement>();
            m_childBlocks = new List<ContentBlock>();
            m_variableReplacer = variableReplacer;
        }

        public IReadOnlyCollection<ContentBlock> ChildBlocks => m_childBlocks;

        public virtual void Expand(ModelDictionary models, OpenXmlElement parentNode)
        {
            var paragraphs = CreateBlockContentForCurrentVariableStack(m_content);
            InsertContent(parentNode, paragraphs);
            ExpandChildBlocks(models, parentNode);
        }

        protected void ExpandChildBlocks(ModelDictionary models, OpenXmlElement parentNode)
        {
            foreach (var child in m_childBlocks)
            {
                child.Expand(models, parentNode);
            }
        }

        protected IEnumerable<OpenXmlElement> CreateBlockContentForCurrentVariableStack(IReadOnlyCollection<OpenXmlElement> content)
        {
            var paragraphs = content.Select(x =>
            {
                var cloned = x.CloneNode(true);
                m_variableReplacer.ReplaceVariables(cloned);
                return cloned;
            });
            return paragraphs;
        }

        protected void InsertContent(OpenXmlElement parentNode, IEnumerable<OpenXmlElement> paragraphs)
        {
            var element = m_leadingPart.GetElement(parentNode);
            if (element == null)
            {
                Console.WriteLine(parentNode.ToPrettyPrintXml());
                throw new OpenXmlTemplateException($"Insertion point {m_leadingPart.Id} not found");
            }
            element.InsertAfterSelf(paragraphs);
        }

        public override string ToString()
        {
            return m_leadingPart?.Id ?? "RootBlock";
        }

        public virtual void SetContent(OpenXmlElement leadingPart, IReadOnlyCollection<OpenXmlElement> loopContent)
        {
            m_leadingPart = InsertionPoint.CreateForElement(leadingPart);
            m_content = loopContent;
        }

        public void AddInnerBlock(ContentBlock block)
        {
            m_childBlocks.Add(block);
        }
    }
}